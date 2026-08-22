import json
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from openpyxl import Workbook

from drive_downloader import (
    DownloadCancelled,
    InventoryWorkbookError,
    _choose_export,
    _csv_safe,
    _validate_resumed_response,
    download_blob,
    download_workspace_file,
    is_selected,
    iter_download_entries,
    safe_destination_path,
    sanitize_path_component,
    scan_inventory_workbook,
)


HEADERS = [
    "Baixar?",
    "Formato de exportação",
    "Tipo",
    "Nome",
    "Tamanho (bytes)",
    "MIME type",
    "ID",
    "Caminho completo",
    "Caminho (JSON)",
    "Resource key",
    "Atalho: ID do alvo",
    "Atalho: MIME type do alvo",
    "Atalho: resource key do alvo",
]


def create_inventory(path: Path, marker="Sim", item_id="file-id") -> None:
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Inventário 1"
    sheet.append(HEADERS)
    sheet.append(
        [
            marker,
            "AUTO",
            "Arquivo",
            "Relatório.pdf",
            42,
            "application/pdf",
            item_id,
            "Raiz/Pasta/Relatório.pdf",
            json.dumps(["Raiz", "Pasta", "Relatório.pdf"]),
            "",
            "",
            "",
            "",
        ]
    )
    workbook.save(path)


class WorkbookSelectionTests(unittest.TestCase):
    def test_selection_markers_are_explicit(self) -> None:
        for marker in ("Sim", "S", "X", "1", True, "TRUE"):
            self.assertTrue(is_selected(marker))
        for marker in ("Não", "", None, False):
            self.assertFalse(is_selected(marker))

    def test_reads_selected_row_and_json_path(self) -> None:
        with tempfile.TemporaryDirectory() as tmpdir:
            path = Path(tmpdir) / "inventory.xlsx"
            create_inventory(path)

            entries = list(iter_download_entries(path))

            self.assertEqual(len(entries), 1)
            self.assertEqual(entries[0].item_id, "file-id")
            self.assertEqual(entries[0].path_segments, ("Raiz", "Pasta", "Relatório.pdf"))

    def test_safe_default_does_not_select_no_rows(self) -> None:
        with tempfile.TemporaryDirectory() as tmpdir:
            path = Path(tmpdir) / "inventory.xlsx"
            create_inventory(path, marker="Não")

            marked = scan_inventory_workbook(path, selected_only=True)
            remaining = scan_inventory_workbook(path, selected_only=False)

            self.assertEqual(marked.candidates, 0)
            self.assertEqual(remaining.candidates, 1)

    def test_requires_expected_headers(self) -> None:
        with tempfile.TemporaryDirectory() as tmpdir:
            path = Path(tmpdir) / "invalid.xlsx"
            workbook = Workbook()
            workbook.active.append(["Coluna qualquer"])
            workbook.save(path)

            with self.assertRaises(InventoryWorkbookError):
                scan_inventory_workbook(path)

    def test_formula_in_id_is_never_interpreted(self) -> None:
        with tempfile.TemporaryDirectory() as tmpdir:
            path = Path(tmpdir) / "inventory.xlsx"
            create_inventory(path, item_id='=WEBSERVICE("https://invalid.example")')

            scan = scan_inventory_workbook(path)

            self.assertEqual(scan.candidates, 0)
            self.assertEqual(scan.invalid_rows, 1)

    def test_csv_report_cells_never_become_formulas(self) -> None:
        self.assertEqual(_csv_safe("=WEBSERVICE('x')"), "'=WEBSERVICE('x')")
        self.assertEqual(_csv_safe("  @SUM(A1)"), "'  @SUM(A1)")
        self.assertEqual(_csv_safe("arquivo.pdf"), "arquivo.pdf")


class SafePathTests(unittest.TestCase):
    def test_windows_reserved_and_invalid_characters_are_neutralized(self) -> None:
        self.assertEqual(sanitize_path_component("CON"), "_CON")
        self.assertEqual(sanitize_path_component("a:b?.txt"), "a_b_.txt")
        self.assertEqual(sanitize_path_component(".."), "sem_nome")

    def test_destination_never_leaves_root_and_collisions_are_renamed(self) -> None:
        with tempfile.TemporaryDirectory() as tmpdir:
            claimed: set[str] = set()
            first = safe_destination_path(
                tmpdir, ("..", "C:\\Windows"), "arquivo.txt", "abcdefgh123", claimed
            )
            second = safe_destination_path(
                tmpdir, ("..", "C:\\Windows"), "arquivo.txt", "zyxwvuts987", claimed
            )
            third = safe_destination_path(
                tmpdir, ("..", "C:\\Windows"), "arquivo.txt", "zyxwvuts987", claimed
            )

            root = Path(tmpdir).resolve()
            self.assertEqual(Path(__import__("os").path.commonpath([root, first])), root)
            self.assertNotEqual(first, second)
            self.assertNotEqual(second, third)
            self.assertIn("zyxwvuts", second.name)


class FakeResponse:
    def __init__(self, content=b"", *, status_code=200, headers=None):
        self.content = content
        self.status_code = status_code
        self.headers = headers or {}
        self.closed = False

    def iter_content(self, chunk_size):
        del chunk_size
        yield self.content

    def raise_for_status(self):
        if self.status_code >= 400:
            raise RuntimeError(f"HTTP {self.status_code}")

    def close(self):
        self.closed = True


class FakeSession:
    def __init__(self, response):
        self.response = response
        self.calls = []
        self.closed = False

    def get(self, url, **kwargs):
        self.calls.append((url, kwargs))
        return self.response

    def close(self):
        self.closed = True


class FakeRequest:
    def __init__(self, response):
        self.response = response
        self.headers = {}

    def execute(self, **_kwargs):
        return self.response


class FakeFiles:
    def __init__(self, operation):
        self.operation = operation
        self.request = None

    def download(self, **kwargs):
        self.kwargs = kwargs
        self.request = FakeRequest(self.operation)
        return self.request


class FakeOperations:
    def get(self, **_kwargs):
        raise AssertionError("uma operação já concluída não deve ser consultada")


class FakeService:
    def __init__(self, operation):
        self.files_resource = FakeFiles(operation)
        self.operations_resource = FakeOperations()

    def files(self):
        return self.files_resource

    def operations(self):
        return self.operations_resource


class DownloadTransferTests(unittest.TestCase):
    def test_blob_resumes_only_from_the_expected_range(self) -> None:
        with tempfile.TemporaryDirectory() as tmpdir:
            destination = Path(tmpdir) / "arquivo.bin"
            part = destination.with_name(".arquivo.bin.driveautomate.part")
            state = part.with_suffix(part.suffix + ".json")
            part.write_bytes(b"abc")
            metadata = {
                "id": "file-id",
                "size": "6",
                "modifiedTime": "2026-01-01T00:00:00Z",
                "version": "7",
            }
            state.write_text(
                json.dumps(
                    {
                        "id": "file-id",
                        "size": "6",
                        "modifiedTime": "2026-01-01T00:00:00Z",
                        "version": "7",
                    }
                ),
                encoding="utf-8",
            )
            response = FakeResponse(
                b"def", status_code=206, headers={"Content-Range": "bytes 3-5/6"}
            )
            session = FakeSession(response)

            with patch("drive_downloader._authorized_session", return_value=session):
                written = download_blob(
                    object(),
                    metadata,
                    destination,
                    latest_metadata_callback=lambda: dict(metadata),
                )

            self.assertEqual(written, 6)
            self.assertEqual(destination.read_bytes(), b"abcdef")
            self.assertEqual(session.calls[0][1]["headers"]["Range"], "bytes=3-")
            self.assertFalse(part.exists())
            self.assertFalse(state.exists())

    def test_invalid_range_is_rejected(self) -> None:
        response = FakeResponse(
            status_code=206, headers={"Content-Range": "bytes 0-3/10"}
        )
        with self.assertRaises(RuntimeError):
            _validate_resumed_response(response, 4, 10)

    def test_workspace_download_uses_completed_lro_and_resource_key(self) -> None:
        operation = {
            "done": True,
            "name": "operations/example",
            "response": {
                "downloadUri": "https://drive.usercontent.google.com/example",
                "partialDownloadAllowed": False,
            },
        }
        service = FakeService(operation)
        session = FakeSession(FakeResponse(b"docx", headers={"Content-Length": "4"}))
        metadata = {"id": "file-id"}
        with tempfile.TemporaryDirectory() as tmpdir, patch(
            "drive_downloader._authorized_session", return_value=session
        ):
            destination = Path(tmpdir) / "documento.docx"
            written = download_workspace_file(
                service,
                object(),
                metadata,
                destination,
                "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                resource_key="resource-key",
            )

            self.assertEqual(written, 4)
            self.assertEqual(destination.read_bytes(), b"docx")
        self.assertEqual(service.files_resource.kwargs["fileId"], "file-id")
        self.assertIn("X-Goog-Drive-Resource-Keys", service.files_resource.request.headers)
        self.assertIn("X-Goog-Drive-Resource-Keys", session.calls[0][1]["headers"])
        self.assertTrue(session.closed)

    def test_workspace_pending_operation_can_be_cancelled_before_polling(self) -> None:
        service = FakeService({"done": False, "name": "operations/example"})
        with tempfile.TemporaryDirectory() as tmpdir:
            with self.assertRaises(DownloadCancelled):
                download_workspace_file(
                    service,
                    object(),
                    {"id": "file-id"},
                    Path(tmpdir) / "documento.docx",
                    "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    cancel_callback=lambda: True,
                )

    def test_vid_auto_format_works_even_when_about_omits_it(self) -> None:
        export_mime, extension = _choose_export(
            "application/vnd.google-apps.vid", "AUTO", {}
        )
        self.assertEqual(export_mime, "video/mp4")
        self.assertEqual(extension, ".mp4")


if __name__ == "__main__":
    unittest.main()
