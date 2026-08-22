"""Configuração local e segura do cliente OAuth do DriveAutomate.

Nenhuma credencial é versionada ou embutida no código-fonte. O responsável pelo
aplicativo importa um JSON OAuth do tipo "Desktop app" no primeiro uso.
"""

from __future__ import annotations

import json
import os
import uuid
from pathlib import Path
from typing import Any

APP_NAME = "DriveAutomate"
OAUTH_CLIENT_FILENAME = "oauth_client.json"


def app_data_dir() -> Path:
    """Diretório privado de dados do usuário atual."""
    appdata = os.getenv("APPDATA")
    if appdata:
        return Path(appdata) / APP_NAME
    return Path.home() / f".{APP_NAME.lower()}"


def oauth_client_path() -> Path:
    return app_data_dir() / OAUTH_CLIENT_FILENAME


def validate_client_config(config: dict[str, Any]) -> dict[str, Any]:
    """Valida e normaliza um JSON OAuth criado para aplicativo desktop."""
    installed = config.get("installed")
    if not isinstance(installed, dict):
        if "web" in config:
            raise ValueError(
                "A credencial selecionada é do tipo Web. Crie um OAuth Client ID "
                "do tipo 'Aplicativo para computador' no Google Cloud."
            )
        raise ValueError("O arquivo não contém uma credencial OAuth de aplicativo desktop.")

    required = ("client_id", "client_secret", "auth_uri", "token_uri")
    missing = [field for field in required if not str(installed.get(field, "")).strip()]
    if missing:
        raise ValueError(
            "O JSON OAuth está incompleto. Campos ausentes: " + ", ".join(missing)
        )

    client_id = str(installed["client_id"]).strip()
    if not client_id.endswith(".apps.googleusercontent.com"):
        raise ValueError("O client_id do arquivo OAuth não parece válido.")

    redirect_uris = installed.get("redirect_uris") or ["http://localhost"]
    if not any(str(uri).startswith("http://localhost") for uri in redirect_uris):
        raise ValueError(
            "A credencial OAuth desktop precisa permitir redirecionamento para localhost."
        )

    return {"installed": dict(installed)}


def load_client_config(path: str | Path | None = None) -> dict[str, Any]:
    source = Path(path) if path else oauth_client_path()
    try:
        raw = json.loads(source.read_text(encoding="utf-8"))
    except FileNotFoundError:
        raise FileNotFoundError(
            "O cliente OAuth ainda não foi configurado. Use 'Configurar OAuth' e "
            "selecione o JSON de um aplicativo desktop."
        ) from None
    except (OSError, json.JSONDecodeError) as exc:
        raise ValueError("Não foi possível ler o arquivo OAuth selecionado.") from exc
    if not isinstance(raw, dict):
        raise ValueError("O arquivo OAuth precisa conter um objeto JSON.")
    return validate_client_config(raw)


def install_client_config(source_path: str | Path) -> Path:
    """Copia uma credencial válida para o perfil local usando troca atômica."""
    config = load_client_config(source_path)
    destination = oauth_client_path()
    destination.parent.mkdir(parents=True, exist_ok=True)
    temporary = destination.with_name(f".{destination.name}.{uuid.uuid4().hex}.tmp")
    try:
        temporary.write_text(
            json.dumps(config, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )
        os.replace(temporary, destination)
    finally:
        if temporary.exists():
            temporary.unlink()
    return destination


def remove_client_config() -> None:
    path = oauth_client_path()
    if path.exists():
        path.unlink()


def has_client_config() -> bool:
    try:
        load_client_config()
    except (OSError, ValueError):
        return False
    return True

