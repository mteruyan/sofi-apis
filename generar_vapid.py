# Genera el par de claves VAPID para los avisos push. Ejecutar UNA vez y guardar los valores como secretos.
import base64
from cryptography.hazmat.primitives.asymmetric import ec
from cryptography.hazmat.primitives import serialization

k = ec.generate_private_key(ec.SECP256R1())
b64 = lambda b: base64.urlsafe_b64encode(b).rstrip(b"=").decode()
print("VAPID_PRIVATE_KEY =", b64(k.private_numbers().private_value.to_bytes(32, "big")))
print("VAPID_PUBLIC_KEY  =", b64(k.public_key().public_bytes(serialization.Encoding.X962,
                                                              serialization.PublicFormat.UncompressedPoint)))
