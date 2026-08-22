import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from drive_inventory import (
    SCOPES,
    api_kwargs,
    format_drive_error,
    format_oauth_error,
    format_user_error,
    is_oauth_client_deleted_error,
    select_oauth_credentials_source,
)
from oauth_config import (
    has_client_config,
    install_client_config,
    load_client_config,
    oauth_client_path,
)


class OAuthErrorHandlingTests(unittest.TestCase):
    def test_detect_deleted_client_error(self) -> None:
        exc = Exception(
            "('deleted_client: The OAuth client was deleted.', "
            "{'error': 'deleted_client', 'error_description': 'The OAuth client was deleted.'})"
        )
        self.assertTrue(is_oauth_client_deleted_error(exc))

    def test_format_deleted_client_error_is_actionable(self) -> None:
        exc = Exception(
            "('deleted_client: The OAuth client was deleted.', "
            "{'error': 'deleted_client', 'error_description': 'The OAuth client was deleted.'})"
        )
        message = format_oauth_error(exc)
        self.assertIn("OAuth Client ID", message)
        self.assertIn("Configurar OAuth", message)

    def test_credentials_file_takes_precedence_over_embedded_config(self) -> None:
        with tempfile.TemporaryDirectory() as tmpdir:
            credentials_path = Path(tmpdir) / "credentials.json"
            credentials_path.write_text("{}", encoding="utf-8")

            resolved_path, resolved_config = select_oauth_credentials_source(
                credentials_path=str(credentials_path),
                client_config={"installed": {"client_id": "abc"}},
            )

            self.assertEqual(resolved_path, str(credentials_path))
            self.assertIsNone(resolved_config)

    def test_format_drive_error_on_unauthorized_link(self) -> None:
        class DummyHttpError(Exception):
            def __init__(self):
                super().__init__("403 Forbidden: File not found")
                self.resp = type("Resp", (), {"status": 403})()

        exc = DummyHttpError()
        message = format_drive_error(exc, "some-folder-id")
        self.assertIn("não foi possível acessar", message.lower())

    def test_format_user_error_for_missing_dependency(self) -> None:
        exc = ModuleNotFoundError("No module named 'google.auth'")
        message = format_user_error(exc)
        self.assertIn("instale as dependências", message.lower())
        self.assertIn("requirements.txt", message)

    def test_format_user_error_for_browser_auth_issue(self) -> None:
        exc = RuntimeError("browser not found for localhost redirect")
        message = format_user_error(exc)
        self.assertIn("janela de login", message.lower())

    def test_api_kwargs_do_not_use_incompatible_shared_drive_argument(self) -> None:
        kwargs = api_kwargs(True)
        self.assertIn("supportsAllDrives", kwargs)
        self.assertNotIn("includeItemsFromAllDrives", kwargs)

    def test_scope_allows_content_download(self) -> None:
        self.assertEqual(SCOPES, ["https://www.googleapis.com/auth/drive.readonly"])

    def test_installs_oauth_only_in_local_app_data(self) -> None:
        config = {
            "installed": {
                "client_id": "example.apps.googleusercontent.com",
                "client_secret": "example-secret",
                "auth_uri": "https://accounts.google.com/o/oauth2/auth",
                "token_uri": "https://oauth2.googleapis.com/token",
                "redirect_uris": ["http://localhost"],
            }
        }
        with tempfile.TemporaryDirectory() as tmpdir, patch.dict(
            "os.environ", {"APPDATA": tmpdir}
        ):
            source = Path(tmpdir) / "source.json"
            source.write_text(__import__("json").dumps(config), encoding="utf-8")

            installed = install_client_config(source)

            self.assertEqual(installed, oauth_client_path())
            self.assertTrue(has_client_config())
            self.assertEqual(
                load_client_config()["installed"]["client_id"],
                "example.apps.googleusercontent.com",
            )


if __name__ == "__main__":
    unittest.main()
