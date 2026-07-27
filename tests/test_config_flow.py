"""Tests for the config flow (config_flow.py)."""

from __future__ import annotations

from custom_components.microsoft_calendar.config_flow import _parse_id_token

from .conftest import make_id_token


# ---------------------------------------------------------------------------
# _parse_id_token
# ---------------------------------------------------------------------------


def test_parse_id_token_extracts_all_claims():
    claims = {
        "oid": "user-object-id-123",
        "name": "Ada Lovelace",
        "preferred_username": "ada@example.com",
        "tid": "tenant-id-456",
    }
    result = _parse_id_token(make_id_token(claims))
    assert result["oid"] == "user-object-id-123"
    assert result["name"] == "Ada Lovelace"
    assert result["preferred_username"] == "ada@example.com"


def test_parse_id_token_not_enough_segments_returns_empty():
    result = _parse_id_token("onlyone")
    assert result == {}


def test_parse_id_token_invalid_base64_returns_empty():
    result = _parse_id_token("header.!!!notbase64!!!.sig")
    assert result == {}


def test_parse_id_token_non_json_payload_returns_empty():
    import base64

    bad_payload = base64.urlsafe_b64encode(b"not json").rstrip(b"=").decode()
    result = _parse_id_token(f"header.{bad_payload}.sig")
    assert result == {}


def test_parse_id_token_extra_padding_handled():
    """Tokens with payloads of varying lengths (padding edge cases) must parse."""
    for extra_claims in [{}, {"x": "y"}, {"a": "b", "c": "d", "e": "f"}]:
        claims = {"oid": "u1", **extra_claims}
        result = _parse_id_token(make_id_token(claims))
        assert result.get("oid") == "u1"


def test_parse_id_token_empty_string_returns_empty():
    result = _parse_id_token("")
    assert result == {}
