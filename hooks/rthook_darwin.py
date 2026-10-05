"""PyInstaller runtime hook — tránh treo/fork trên macOS."""
import os
import sys

if sys.platform == 'darwin':
    os.environ.setdefault('OBJC_DISABLE_INITIALIZE_FORK_SAFETY', 'YES')
    # Giảm prompt Keychain chặn luồng chính lúc mở app đóng gói.
    # Phải là đường dẫn tới CLASS: keyring tách ở dấu chấm cuối rồi getattr,
    # trỏ tới module thì get_keyring() ném AttributeError.
    if getattr(sys, 'frozen', False):
        os.environ.setdefault('PYTHON_KEYRING_BACKEND', 'keyring.backends.macOS.Keyring')
