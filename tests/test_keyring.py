"""P0.1 — a chave da API nunca vai para backend inseguro nem para o config."""
import pytest

from conftest import et


def _fake_backend(module: str, priority: float):
    cls = type("FakeKeyring", (), {"priority": priority})
    cls.__module__ = module
    return cls()


@pytest.mark.parametrize("module,secure", [
    ("keyring.backends.SecretService", True),
    ("keyring.backends.macOS", True),
    ("keyrings.alt.file", False),            # PlaintextKeyring / EncryptedKeyring
    ("keyring.backends.fail", False),
    ("keyring.backends.null", False),
    ("keyring.backends.chainer", False),     # pode cair no PlaintextKeyring
])
def test_backend_classification(module, secure):
    assert et._is_secure_backend(_fake_backend(module, 1)) is secure


def test_select_prefers_secure_even_if_plaintext_present(monkeypatch):
    import keyring
    import keyring.backend
    secure = _fake_backend("keyring.backends.SecretService", 5)
    plain = _fake_backend("keyrings.alt.file", 0.5)
    chosen = []
    monkeypatch.setattr(keyring.backend, "get_all_keyring", lambda: [plain, secure])
    monkeypatch.setattr(keyring, "set_keyring", chosen.append)
    assert et._select_secure_keyring() is secure
    assert chosen == [secure]


def test_select_returns_none_when_only_insecure(monkeypatch):
    import keyring
    import keyring.backend
    plain = _fake_backend("keyrings.alt.file", 0.5)
    chosen = []
    monkeypatch.setattr(keyring.backend, "get_all_keyring", lambda: [plain])
    monkeypatch.setattr(keyring, "set_keyring", chosen.append)
    assert et._select_secure_keyring() is None
    assert chosen == []          # nunca instala o backend inseguro


def test_save_refuses_without_secure_keyring(monkeypatch):
    monkeypatch.setattr(et, "HAS_KEYRING", False)
    assert et.keyring_save_key("sk-ant-teste") is False


def test_default_config_never_contains_key():
    assert et._CONFIG_DEFAULTS["ai_key"] == ""


def test_env_var_status_message(monkeypatch):
    monkeypatch.setenv(et.API_KEY_ENV, "sk-ant-teste")
    assert et.API_KEY_ENV in et.api_key_storage_status()
