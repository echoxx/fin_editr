#!/usr/bin/env python3
"""
Secure credential management for Investing.com Pro login.

Storage priority:
1. Environment variables (INVESTING_COM_EMAIL, INVESTING_COM_PASSWORD)
2. OS keyring (Windows Credential Manager, macOS Keychain, Linux Secret Service)
3. Encrypted local file (~/.fin_editr/credentials.json) - fallback for WSL

The Gmail login note: Investing.com allows login via Google OAuth, but for
automation we need the direct email/password. If you normally use "Sign in
with Google", you may need to set a password on your Investing.com account.
"""

import os
import json
import base64
import getpass
import sys
from pathlib import Path

# Try to import keyring
try:
    import keyring
    from keyring.errors import NoKeyringError
    KEYRING_AVAILABLE = True
except ImportError:
    KEYRING_AVAILABLE = False
    NoKeyringError = Exception  # Dummy for except clause

# Try to import cryptography for secure file storage
try:
    from cryptography.fernet import Fernet
    from cryptography.hazmat.primitives import hashes
    from cryptography.hazmat.primitives.kdf.pbkdf2 import PBKDF2HMAC
    CRYPTO_AVAILABLE = True
except ImportError:
    CRYPTO_AVAILABLE = False

# Constants
SERVICE_NAME = "fin_editr_investing_com"
USERNAME_KEY = "email"
CREDENTIALS_DIR = Path.home() / ".fin_editr"
CREDENTIALS_FILE = CREDENTIALS_DIR / "credentials.enc"
SALT_FILE = CREDENTIALS_DIR / ".salt"


def _get_machine_key() -> bytes:
    """Generate a machine-specific key for file encryption."""
    # Use a combination of username and machine identifiers
    import socket
    machine_id = f"{os.getlogin()}@{socket.gethostname()}"
    return machine_id.encode()


def _derive_key(salt: bytes) -> bytes:
    """Derive an encryption key from machine-specific data."""
    if not CRYPTO_AVAILABLE:
        return base64.b64encode(_get_machine_key())[:32]

    kdf = PBKDF2HMAC(
        algorithm=hashes.SHA256(),
        length=32,
        salt=salt,
        iterations=100000,
    )
    return base64.urlsafe_b64encode(kdf.derive(_get_machine_key()))


def _save_to_file(email: str, password: str) -> bool:
    """Save credentials to encrypted local file."""
    try:
        CREDENTIALS_DIR.mkdir(parents=True, exist_ok=True)

        # Generate or load salt
        if SALT_FILE.exists():
            salt = SALT_FILE.read_bytes()
        else:
            salt = os.urandom(16)
            SALT_FILE.write_bytes(salt)
            # Make salt file readable only by owner
            SALT_FILE.chmod(0o600)

        key = _derive_key(salt)
        data = json.dumps({"email": email, "password": password})

        if CRYPTO_AVAILABLE:
            f = Fernet(key)
            encrypted = f.encrypt(data.encode())
        else:
            # Fallback: base64 encoding (obfuscation, not true encryption)
            encrypted = base64.b64encode(data.encode())

        CREDENTIALS_FILE.write_bytes(encrypted)
        CREDENTIALS_FILE.chmod(0o600)  # Readable only by owner

        return True
    except Exception as e:
        print(f"Error saving to file: {e}")
        return False


def _load_from_file() -> tuple[str, str] | None:
    """Load credentials from encrypted local file."""
    try:
        if not CREDENTIALS_FILE.exists():
            return None

        if not SALT_FILE.exists():
            return None

        salt = SALT_FILE.read_bytes()
        key = _derive_key(salt)
        encrypted = CREDENTIALS_FILE.read_bytes()

        if CRYPTO_AVAILABLE:
            f = Fernet(key)
            data = json.loads(f.decrypt(encrypted).decode())
        else:
            # Fallback: base64 decoding
            data = json.loads(base64.b64decode(encrypted).decode())

        return data.get("email"), data.get("password")
    except Exception:
        return None


def _delete_file_credentials() -> bool:
    """Delete file-based credentials."""
    try:
        if CREDENTIALS_FILE.exists():
            CREDENTIALS_FILE.unlink()
        if SALT_FILE.exists():
            SALT_FILE.unlink()
        return True
    except Exception:
        return False


def _try_keyring_set(service: str, key: str, value: str) -> bool:
    """Try to set a keyring value, return False if keyring unavailable."""
    if not KEYRING_AVAILABLE:
        return False
    try:
        keyring.set_password(service, key, value)
        return True
    except (NoKeyringError, Exception) as e:
        if "No recommended backend" in str(e) or isinstance(e, NoKeyringError):
            return False
        raise


def _try_keyring_get(service: str, key: str) -> str | None:
    """Try to get a keyring value, return None if keyring unavailable."""
    if not KEYRING_AVAILABLE:
        return None
    try:
        return keyring.get_password(service, key)
    except (NoKeyringError, Exception):
        return None


def credentials_exist() -> bool:
    """Check if credentials are configured."""
    # Check environment variables first
    if os.environ.get("INVESTING_COM_EMAIL") and os.environ.get("INVESTING_COM_PASSWORD"):
        return True

    # Check keyring
    email = _try_keyring_get(SERVICE_NAME, USERNAME_KEY)
    if email:
        password = _try_keyring_get(SERVICE_NAME, email)
        if password:
            return True

    # Check file storage
    result = _load_from_file()
    if result and result[0] and result[1]:
        return True

    return False


def get_credentials() -> tuple[str, str]:
    """
    Retrieve stored credentials.

    Returns:
        Tuple of (email, password)

    Raises:
        ValueError: If no credentials are configured
    """
    # Check environment variables first (highest priority)
    env_email = os.environ.get("INVESTING_COM_EMAIL")
    env_password = os.environ.get("INVESTING_COM_PASSWORD")

    if env_email and env_password:
        return env_email, env_password

    # Check keyring
    email = _try_keyring_get(SERVICE_NAME, USERNAME_KEY)
    if email:
        password = _try_keyring_get(SERVICE_NAME, email)
        if password:
            return email, password

    # Check file storage
    result = _load_from_file()
    if result and result[0] and result[1]:
        return result

    raise ValueError(
        "No credentials configured. Run 'python netnet_main.py setup-credentials' "
        "or set INVESTING_COM_EMAIL and INVESTING_COM_PASSWORD environment variables."
    )


def setup_credentials() -> bool:
    """
    Interactive setup - prompts for credentials and stores securely.

    Returns:
        True if credentials were stored successfully
    """
    print("=" * 50)
    print("Investing.com Pro Credential Setup")
    print("=" * 50)

    print("\nNOTE: If you normally use 'Sign in with Google' on Investing.com,")
    print("you may need to set a password on your account first at:")
    print("https://www.investing.com/members-admin/account\n")

    # Get email
    email = input("Email address: ").strip()
    if not email:
        print("ERROR: Email cannot be empty.")
        return False

    # Get password (hidden input)
    password = getpass.getpass("Password: ")
    if not password:
        print("ERROR: Password cannot be empty.")
        return False

    # Confirm password
    password_confirm = getpass.getpass("Confirm password: ")
    if password != password_confirm:
        print("ERROR: Passwords do not match.")
        return False

    # Try keyring first
    keyring_success = False
    try:
        if _try_keyring_set(SERVICE_NAME, USERNAME_KEY, email):
            if _try_keyring_set(SERVICE_NAME, email, password):
                keyring_success = True
                print("\nCredentials stored in OS keyring.")
    except Exception:
        pass

    # Fall back to file storage if keyring didn't work
    if not keyring_success:
        print("\nOS keyring not available (common on WSL).")
        print(f"Storing credentials in: {CREDENTIALS_FILE}")

        if not CRYPTO_AVAILABLE:
            print("WARNING: 'cryptography' package not installed.")
            print("         Credentials will be obfuscated but not fully encrypted.")
            print("         For better security: pip install cryptography")

        if not _save_to_file(email, password):
            print("\nERROR: Failed to store credentials.")
            return False

    print("\nCredentials stored successfully!")
    print(f"Email: {email}")
    print("Password: ********")
    return True


def delete_credentials() -> bool:
    """
    Remove stored credentials.

    Returns:
        True if credentials were deleted successfully
    """
    deleted_keyring = False
    deleted_file = False

    # Try deleting from keyring
    if KEYRING_AVAILABLE:
        try:
            email = _try_keyring_get(SERVICE_NAME, USERNAME_KEY)
            if email:
                try:
                    keyring.delete_password(SERVICE_NAME, email)
                    deleted_keyring = True
                except Exception:
                    pass
                try:
                    keyring.delete_password(SERVICE_NAME, USERNAME_KEY)
                    deleted_keyring = True
                except Exception:
                    pass
        except Exception:
            pass

    # Delete file credentials
    if _delete_file_credentials():
        deleted_file = True

    if deleted_keyring or deleted_file:
        print("Credentials deleted successfully.")
        return True
    else:
        print("No credentials found to delete.")
        return True


def verify_credentials() -> bool:
    """
    Verify that credentials exist and can be retrieved.
    Does NOT verify they are valid with Investing.com.

    Returns:
        True if credentials can be retrieved
    """
    try:
        email, password = get_credentials()

        # Mask password for display
        if len(password) > 2:
            masked_password = password[0] + "*" * (len(password) - 2) + password[-1]
        else:
            masked_password = "**"

        print("Credentials found:")
        print(f"  Email: {email}")
        print(f"  Password: {masked_password}")

        # Determine source
        if os.environ.get("INVESTING_COM_EMAIL"):
            print("  Source: Environment variables")
        elif _try_keyring_get(SERVICE_NAME, USERNAME_KEY):
            print("  Source: OS Keyring")
        elif CREDENTIALS_FILE.exists():
            print(f"  Source: Encrypted file ({CREDENTIALS_FILE})")
        else:
            print("  Source: Unknown")

        return True

    except ValueError as e:
        print(f"ERROR: {e}")
        return False


def main():
    """CLI for credential management."""
    import argparse

    parser = argparse.ArgumentParser(description="Manage Investing.com credentials")
    parser.add_argument("action", choices=["setup", "verify", "delete"],
                        help="Action to perform")

    args = parser.parse_args()

    if args.action == "setup":
        success = setup_credentials()
        sys.exit(0 if success else 1)
    elif args.action == "verify":
        success = verify_credentials()
        sys.exit(0 if success else 1)
    elif args.action == "delete":
        success = delete_credentials()
        sys.exit(0 if success else 1)


if __name__ == "__main__":
    main()
