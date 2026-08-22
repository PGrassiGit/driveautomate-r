import json
import os
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from drive_inventory import (
    SCOPES,
    account_token_path,
    authenticate,
    sanitize_account_filename,
    token_identity,
)
from secure_store import is_protected_file, read_secure_json, write_secure_json


class SecureStoreTests(unittest.TestCase):
    def test_round_trip_and_windows_protection(self) -> None:
        with tempfile.TemporaryDirectory() as tmpdir:
            path = Path(tmpdir) / "token.dat"
            payload = {"credentials": {"token": "example"}, "account_email": "a@example.test"}

            write_secure_json(path, payload)

            self.assertEqual(read_secure_json(path), payload)
            if os.name == "nt":
                self.assertTrue(is_protected_file(path))
                self.assertNotIn(b"example", path.read_bytes())

    def test_account_filename_is_stable_and_does_not_expose_email(self) -> None:
        first = sanitize_account_filename("Person@Example.test")
        second = sanitize_account_filename("person@example.test")
        self.assertEqual(first, second)
        self.assertNotIn("person", first)
        self.assertNotIn("@", first)

        with tempfile.TemporaryDirectory() as tmpdir, patch.dict(
            os.environ, {"APPDATA": tmpdir}
        ):
            path = account_token_path("person@example.test")
            self.assertEqual(path.suffix, ".dat")
            self.assertNotIn("person", path.name)

    def test_legacy_authorized_json_is_migrated_without_new_login(self) -> None:
        credentials = {
            "token": "access-example",
            "refresh_" + "token": "refresh-example",
            "token_uri": "https://oauth2.googleapis.com/token",
            "client_id": "example.apps.googleusercontent.com",
            "client_secret": "example-secret",
            "scopes": SCOPES,
            "expiry": "2999-01-01T00:00:00Z",
        }
        with tempfile.TemporaryDirectory() as tmpdir:
            path = Path(tmpdir) / "legacy.json"
            path.write_text(json.dumps(credentials), encoding="utf-8")

            result = authenticate(None, path)

            self.assertTrue(result.valid)
            stored = read_secure_json(path)
            self.assertEqual(stored["credentials"]["scopes"], SCOPES)
            if os.name == "nt":
                self.assertTrue(is_protected_file(path))

    def test_identity_is_read_from_protected_envelope(self) -> None:
        with tempfile.TemporaryDirectory() as tmpdir:
            path = Path(tmpdir) / "account.dat"
            write_secure_json(
                path,
                {
                    "schema": 1,
                    "credentials": {},
                    "account_email": "person@example.test",
                    "account_display_name": "Example Person",
                },
            )
            self.assertEqual(
                token_identity(path), ("person@example.test", "Example Person")
            )


if __name__ == "__main__":
    unittest.main()
