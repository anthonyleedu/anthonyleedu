from src.llm_client import AuthenticationFailed, is_auth_error, sanitize_error


def test_sanitize_error_strips_openai_key_material():
    raw = "Error code: 401 - Incorrect API key provided: sk-proj-abcDEF123-restofkeyazEA"
    cleaned = sanitize_error(Exception(raw))
    assert "sk-proj-abcDEF123" not in cleaned
    assert "restofkeyazEA" not in cleaned
    assert "sk-***" in cleaned


def test_is_auth_error_detects_invalidated_and_incorrect_keys():
    assert is_auth_error(Exception("Error code: 401 - token_invalidated"))
    assert is_auth_error(Exception("Incorrect API key provided: sk-***azEA"))
    assert is_auth_error(AuthenticationFailed("rejected"))
    assert not is_auth_error(Exception("rate limit exceeded"))
