#!/usr/bin/env python3
"""
Secure credential management for Investing.com Pro login.

Uses OS-native secure storage via keyring library:
- Windows: Windows Credential Manager
- macOS: Keychain
- Linux: Secret Service (GNOME Keyring, KWallet)

Supports environment variable override for CI/automation.
"""

import os
import getpass
import sys

try:
    import keyring
    KEYRING_AVAILABLE = True
except ImportError:
    KEYRING_AVAILABLE = False

# Service identifier for keyring storage
SERVICE_NAME = "fin_editr_investing_com"
USERNAME_KEY = "email"


def credentials_exist() -> bool:
    """Check if credentials are configured (keyring or environment)."""
    # Check environment variables first
    if os.environ.get("INVESTING_COM_EMAIL") and os.environ.get("INVESTING_COM_PASSWORD"):
        return True

    # Check keyring
    if KEYRING_AVAILABLE:
        try:
            email = keyring.get_password(SERVICE_NAME, USERNAME_KEY)
            if email:
                password = keyring.get_password(SERVICE_NAME, email)
                return password is not None
        except Exception:
            pass

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
    if KEYRING_AVAILABLE:
        try:
            email = keyring.get_password(SERVICE_NAME, USERNAME_KEY)
            if email:
                password = keyring.get_password(SERVICE_NAME, email)
                if password:
                    return email, password
        except Exception as e:
            raise ValueError(f"Error retrieving credentials from keyring: {e}")

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
    if not KEYRING_AVAILABLE:
        print("ERROR: keyring library not installed.")
        print("Install with: pip install keyring")
        print("\nAlternatively, set environment variables:")
        print("  INVESTING_COM_EMAIL=your@email.com")
        print("  INVESTING_COM_PASSWORD=yourpassword")
        return False

    print("=" * 50)
    print("Investing.com Pro Credential Setup")
    print("=" * 50)
    print("\nCredentials will be stored securely in your OS keyring.")
    print("(Windows Credential Manager / macOS Keychain / Linux Secret Service)\n")

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

    try:
        # Store email reference
        keyring.set_password(SERVICE_NAME, USERNAME_KEY, email)
        # Store password under email key
        keyring.set_password(SERVICE_NAME, email, password)

        print("\nCredentials stored successfully!")
        print(f"Email: {email}")
        print("Password: ********")
        return True

    except Exception as e:
        print(f"\nERROR: Failed to store credentials: {e}")
        return False


def delete_credentials() -> bool:
    """
    Remove stored credentials from keyring.

    Returns:
        True if credentials were deleted successfully
    """
    if not KEYRING_AVAILABLE:
        print("ERROR: keyring library not installed.")
        return False

    try:
        # Get current email to delete its password entry
        email = keyring.get_password(SERVICE_NAME, USERNAME_KEY)

        if email:
            # Delete password
            try:
                keyring.delete_password(SERVICE_NAME, email)
            except keyring.errors.PasswordDeleteError:
                pass  # Already deleted or doesn't exist

            # Delete email reference
            try:
                keyring.delete_password(SERVICE_NAME, USERNAME_KEY)
            except keyring.errors.PasswordDeleteError:
                pass

            print("Credentials deleted successfully.")
            return True
        else:
            print("No credentials found to delete.")
            return True

    except Exception as e:
        print(f"ERROR: Failed to delete credentials: {e}")
        return False


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
        masked_password = password[0] + "*" * (len(password) - 2) + password[-1] if len(password) > 2 else "**"

        print("Credentials found:")
        print(f"  Email: {email}")
        print(f"  Password: {masked_password}")

        # Check source
        if os.environ.get("INVESTING_COM_EMAIL"):
            print("  Source: Environment variables")
        else:
            print("  Source: OS Keyring")

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
