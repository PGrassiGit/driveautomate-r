"""Armazenamento local de JSON sensível, protegido pelo usuário do Windows."""

from __future__ import annotations

import ctypes
import json
import os
import uuid
from ctypes import wintypes
from pathlib import Path
from typing import Any

MAGIC = b"DriveAutomate-DPAPI-v1\0"
CRYPTPROTECT_UI_FORBIDDEN = 0x01


class _DataBlob(ctypes.Structure):
    _fields_ = [("cbData", wintypes.DWORD), ("pbData", ctypes.POINTER(ctypes.c_byte))]


def _blob(data: bytes) -> tuple[_DataBlob, ctypes.Array]:
    buffer = (ctypes.c_byte * len(data)).from_buffer_copy(data)
    return _DataBlob(len(data), ctypes.cast(buffer, ctypes.POINTER(ctypes.c_byte))), buffer


def _windows_apis():
    crypt32 = ctypes.windll.crypt32
    kernel32 = ctypes.windll.kernel32
    crypt32.CryptProtectData.argtypes = [
        ctypes.POINTER(_DataBlob),
        wintypes.LPCWSTR,
        ctypes.POINTER(_DataBlob),
        ctypes.c_void_p,
        ctypes.c_void_p,
        wintypes.DWORD,
        ctypes.POINTER(_DataBlob),
    ]
    crypt32.CryptProtectData.restype = wintypes.BOOL
    crypt32.CryptUnprotectData.argtypes = [
        ctypes.POINTER(_DataBlob),
        ctypes.POINTER(wintypes.LPWSTR),
        ctypes.POINTER(_DataBlob),
        ctypes.c_void_p,
        ctypes.c_void_p,
        wintypes.DWORD,
        ctypes.POINTER(_DataBlob),
    ]
    crypt32.CryptUnprotectData.restype = wintypes.BOOL
    kernel32.LocalFree.argtypes = [ctypes.c_void_p]
    kernel32.LocalFree.restype = ctypes.c_void_p
    return crypt32, kernel32


def _protect_windows(data: bytes) -> bytes:
    input_blob, input_buffer = _blob(data)
    output_blob = _DataBlob()
    crypt32, kernel32 = _windows_apis()
    if not crypt32.CryptProtectData(
        ctypes.byref(input_blob),
        "DriveAutomate OAuth token",
        None,
        None,
        None,
        CRYPTPROTECT_UI_FORBIDDEN,
        ctypes.byref(output_blob),
    ):
        raise OSError(ctypes.get_last_error(), "O Windows não conseguiu proteger o token.")
    try:
        return ctypes.string_at(output_blob.pbData, output_blob.cbData)
    finally:
        del input_buffer
        kernel32.LocalFree(output_blob.pbData)


def _unprotect_windows(data: bytes) -> bytes:
    input_blob, input_buffer = _blob(data)
    output_blob = _DataBlob()
    description = wintypes.LPWSTR()
    crypt32, kernel32 = _windows_apis()
    if not crypt32.CryptUnprotectData(
        ctypes.byref(input_blob),
        ctypes.byref(description),
        None,
        None,
        None,
        CRYPTPROTECT_UI_FORBIDDEN,
        ctypes.byref(output_blob),
    ):
        raise OSError(
            ctypes.get_last_error(),
            "O token não pertence a este usuário do Windows ou está corrompido.",
        )
    try:
        return ctypes.string_at(output_blob.pbData, output_blob.cbData)
    finally:
        del input_buffer
        kernel32.LocalFree(output_blob.pbData)
        if description:
            kernel32.LocalFree(description)


def is_protected_file(path: str | Path) -> bool:
    try:
        with Path(path).open("rb") as stream:
            return stream.read(len(MAGIC)) == MAGIC
    except OSError:
        return False


def read_secure_json(path: str | Path) -> dict[str, Any]:
    raw = Path(path).read_bytes()
    if raw.startswith(MAGIC):
        if os.name != "nt":
            raise OSError("Este token protegido por DPAPI só pode ser aberto no Windows.")
        raw = _unprotect_windows(raw[len(MAGIC) :])
    try:
        result = json.loads(raw.decode("utf-8"))
    except (UnicodeDecodeError, json.JSONDecodeError) as exc:
        raise ValueError("O arquivo de token está corrompido ou é inválido.") from exc
    if not isinstance(result, dict):
        raise ValueError("O arquivo de token precisa conter um objeto JSON.")
    return result


def write_secure_json(path: str | Path, payload: dict[str, Any]) -> None:
    destination = Path(path)
    destination.parent.mkdir(parents=True, exist_ok=True)
    raw = json.dumps(payload, ensure_ascii=False, separators=(",", ":")).encode("utf-8")
    if os.name == "nt":
        raw = MAGIC + _protect_windows(raw)
    temporary = destination.with_name(f".{destination.name}.{uuid.uuid4().hex}.tmp")
    try:
        temporary.write_bytes(raw)
        if os.name != "nt":
            temporary.chmod(0o600)
        os.replace(temporary, destination)
    finally:
        temporary.unlink(missing_ok=True)
