"""Import this BEFORE ios_xe_upgrade_manager in every test module.

It points the storage layer at a throw-away folder, so tests never read or change the real database, the real
vault (or its key in Windows Credential Manager), or the real storage folder with the sites' backups.
"""
import atexit
import os
import shutil
import tempfile

ROOT = tempfile.mkdtemp(prefix="netupgrade-tests-")
os.environ["NETUPGRADE_DATA_DIR"] = os.path.join(ROOT, "data")
os.environ["NETUPGRADE_STORAGE_DIR"] = os.path.join(ROOT, "storage")
os.environ["NETUPGRADE_VAULT_KEY"] = "memory"
atexit.register(shutil.rmtree, ROOT, True)
