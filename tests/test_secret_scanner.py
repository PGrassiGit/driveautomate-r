from scripts.check_no_secrets import PATTERNS


def test_google_api_key_pattern_detects_realistic_value():
    value = ("AI" + "za" + "A" * 35).encode()

    assert PATTERNS["Google API key"].search(value)


def test_google_oauth_client_id_pattern_detects_realistic_value():
    value = ("123456-" + "a" * 20 + ".apps.googleusercontent.com").encode()

    assert PATTERNS["Google OAuth client ID real"].search(value)


def test_public_oauth_placeholders_are_allowed():
    example = (
        b'"client_id": "YOUR_CLIENT_ID.apps.googleusercontent.com", '
        b'"client_secret": "YOUR_CLIENT_SECRET"'
    )

    assert not any(pattern.search(example) for pattern in PATTERNS.values())
