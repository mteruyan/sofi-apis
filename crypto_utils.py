"""
crypto_utils.py
----------------
Cifrado a nivel de aplicación para datos sensibles de salud (Ley 29733).
Usa AES-256-GCM. La llave se carga UNA VEZ por invocación (patrón cold-start,
igual que las credenciales de Danaconnect) desde Secret Manager.

Uso:
    from crypto_utils import cifrar, descifrar, cifrar_json, descifrar_json

    valor_enc = cifrar("75")            # -> bytes listos para BYTEA
    valor     = descifrar(valor_enc)    # -> "75"

    resp_enc = cifrar_json({"1.2": "N", "1.3": "S"})
    resp     = descifrar_json(resp_enc)
"""

import os
import json
import base64
from google.cloud import secretmanager

_AESGCM = None  # cache a nivel de módulo (cold start)


def _get_secret(secret_id: str, project_id: str = None) -> str:
    project_id = project_id or os.environ.get("GCP_PROJECT", "sanitas-recetas-prod")
    client = secretmanager.SecretManagerServiceClient()
    name = f"projects/{project_id}/secrets/{secret_id}/versions/latest"
    response = client.access_secret_version(request={"name": name})
    return response.payload.data.decode("UTF-8")


def _get_cipher():
    """
    Carga (y cachea) la llave AES-256 desde Secret Manager.
    El secreto 'afiliacion-salud-key' debe contener 32 bytes en base64
    (generar una sola vez con: openssl rand -base64 32).
    """
    global _AESGCM
    if _AESGCM is None:
        from cryptography.hazmat.primitives.ciphers.aead import AESGCM
        key_b64 = _get_secret("afiliacion-salud-key")
        key = base64.b64decode(key_b64)
        if len(key) != 32:
            raise ValueError("La llave 'afiliacion-salud-key' debe ser de 32 bytes (AES-256)")
        _AESGCM = AESGCM(key)
    return _AESGCM


def cifrar(texto: str) -> bytes:
    """
    Cifra un string y devuelve bytes listos para insertar en columna BYTEA.
    Formato almacenado: nonce(12 bytes) + ciphertext (incluye tag GCM).
    """
    if texto is None:
        return None
    aesgcm = _get_cipher()
    nonce = os.urandom(12)
    ct = aesgcm.encrypt(nonce, texto.encode("utf-8"), None)
    return nonce + ct


def descifrar(data: bytes) -> str:
    """Descifra bytes almacenados en BYTEA y devuelve el string original."""
    if data is None:
        return None
    aesgcm = _get_cipher()
    nonce, ct = data[:12], data[12:]
    pt = aesgcm.decrypt(nonce, ct, None)
    return pt.decode("utf-8")


def cifrar_json(obj: dict) -> bytes:
    """Serializa un dict a JSON y lo cifra."""
    if obj is None:
        return None
    return cifrar(json.dumps(obj, ensure_ascii=False))


def descifrar_json(data: bytes) -> dict:
    """Descifra y deserializa un JSON previamente cifrado con cifrar_json()."""
    if data is None:
        return None
    return json.loads(descifrar(data))
