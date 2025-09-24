import os
import base64
from cryptography.fernet import Fernet
from cryptography.hazmat.primitives import hashes
from cryptography.hazmat.primitives.kdf.pbkdf2 import PBKDF2HMAC
from flask import current_app

def get_key_from_password(salt):
    """Derives a Fernet key from the app's SECRET_KEY and a provided salt."""
    password = current_app.config['SECRET_KEY'].encode()
    kdf = PBKDF2HMAC(
        algorithm=hashes.SHA256(),
        length=32,
        salt=salt,
        iterations=100000,
    )
    key = base64.urlsafe_b64encode(kdf.derive(password))
    return key

def encrypt_password(password_to_encrypt):
    """
    Encrypts a password using a randomly generated salt.
    Returns a base64 encoded string containing both the salt and the ciphertext.
    """
    if not password_to_encrypt:
        return ""

    salt = os.urandom(16)
    key = get_key_from_password(salt)
    f = Fernet(key)

    encrypted_password = f.encrypt(password_to_encrypt.encode())

    # Combine salt and encrypted password, then encode to base64 to get a storable string
    storable_password = base64.urlsafe_b64encode(salt + encrypted_password)

    return storable_password.decode('utf-8')

def decrypt_password(storable_password_b64):
    """
    Decrypts a password from a storable base64 string that contains the salt.
    """
    if not storable_password_b64:
        return ""

    try:
        # Decode the base64 string to get the raw bytes
        decoded_storable_password = base64.urlsafe_b64decode(storable_password_b64)

        # Extract the salt (first 16 bytes) and the actual encrypted password
        salt = decoded_storable_password[:16]
        encrypted_password = decoded_storable_password[16:]

        key = get_key_from_password(salt)
        f = Fernet(key)

        decrypted_password = f.decrypt(encrypted_password)
        return decrypted_password.decode('utf-8')
    except Exception:
        # If decryption fails for any reason (e.g., wrong key, corrupted data),
        # return an empty string or handle the error as appropriate.
        # Avoid returning the encrypted value.
        return ""
