"""Falha o build quando encontra credenciais ou caminhos pessoais no repositório."""

from __future__ import annotations

import re
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
IGNORED_PARTS = {".git", ".venv", "__pycache__", ".pytest_cache"}
MAX_FILE_SIZE = 512 * 1024 * 1024
FORBIDDEN_DIRECTORIES = {"build", "dist"}
FORBIDDEN_ARCHIVE_SUFFIXES = {".zip", ".7z", ".rar"}

# As expressões são montadas em partes para o scanner não acusar a própria fonte.
PATTERNS = {
    "Google OAuth client secret": re.compile((b"GOC" + b"SPX-[A-Za-z0-9_-]{20,}")),
    "Google OAuth client ID real": re.compile(
        b"[0-9]{6,}-[a-z0-9_-]{20,}\\.apps\\.googleusercontent\\.com",
        re.IGNORECASE,
    ),
    "refresh token OAuth": re.compile(b'"refresh_' + b'token"\\s*:', re.IGNORECASE),
    "private key": re.compile(b"-----BEGIN (?:RSA |EC |)PRIVATE KEY-----"),
    "caminho de perfil Windows": re.compile(
        b"[A-Za-z]:[\\\\/]Users[\\\\/](?!Public(?:[\\\\/]|$))[^\\\\/\\r\\n]+",
        re.IGNORECASE,
    ),
}


def iter_files():
    for path in ROOT.rglob("*"):
        if not path.is_file() or any(part in IGNORED_PARTS for part in path.parts):
            continue
        if path.resolve() == Path(__file__).resolve():
            continue
        try:
            if path.stat().st_size <= MAX_FILE_SIZE:
                yield path
        except OSError:
            continue


def main() -> int:
    findings: list[str] = []
    for path in iter_files():
        relative = path.relative_to(ROOT)
        if any(part.casefold() in FORBIDDEN_DIRECTORIES for part in relative.parts[:-1]):
            findings.append(f"{relative}: artefato de build não permitido")
            continue
        if path.suffix.casefold() in FORBIDDEN_ARCHIVE_SUFFIXES:
            findings.append(f"{relative}: arquivo compactado não permitido")
            continue
        if (
            path.name.casefold().startswith(("credentials", "token"))
            and path.suffix.casefold() == ".json"
            and path.name.casefold() != "credentials.example.json"
        ):
            findings.append(f"{relative}: JSON de credencial/token não permitido")
            continue
        try:
            content = path.read_bytes()
        except OSError:
            continue
        for label, pattern in PATTERNS.items():
            if pattern.search(content):
                findings.append(f"{relative}: {label}")

    if findings:
        print("Publicação bloqueada: conteúdo sensível encontrado:")
        for finding in sorted(findings):
            print(f"- {finding}")
        return 1

    print("Verificação concluída: nenhum segredo ou caminho pessoal conhecido encontrado.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
