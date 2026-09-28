"""
Encrypts the generated .xlsx with a file-open password.

openpyxl cannot write encrypted workbooks, so we:
  1. Save the plain workbook to a temp file.
  2. Use msoffcrypto-tool to produce an encrypted copy.
  3. Replace the plain file with the encrypted one.
  4. Delete the temp file.
"""
import io
import os
import tempfile
from pathlib import Path

import msoffcrypto


def encrypt_xlsx(plain_path: str, password: str,
                 encrypted_path: str | None = None) -> str:
    """
    Encrypt an .xlsx file with the given password.

    Parameters
    ----------
    plain_path      : source .xlsx (unencrypted)
    password        : file-open password
    encrypted_path  : destination; defaults to <plain_path> (in-place)

    Returns
    -------
    Path to the encrypted file.
    """
    plain = Path(plain_path)
    target = Path(encrypted_path) if encrypted_path else plain

    if not plain.exists():
        raise FileNotFoundError(plain)

    # Write encrypted output to a temp file first, then swap
    tmp_fd, tmp_name = tempfile.mkstemp(suffix=".xlsx")
    os.close(tmp_fd)

    try:
        with open(plain, "rb") as f_in:
            office_file = msoffcrypto.OfficeFile(f_in)
            # 2007+ format supports AES encryption
            office_file.load_key(password=password, verify_password=True)
            with open(tmp_name, "wb") as f_out:
                office_file.encrypt(password, f_out)

        # Atomic replace
        os.replace(tmp_name, target)
    except Exception:
        # Clean up temp file on failure
        if os.path.exists(tmp_name):
            os.remove(tmp_name)
        raise

    return str(target)