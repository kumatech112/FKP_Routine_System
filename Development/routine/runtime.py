"""Restore absent standard Windows process variables from native known folders."""
import os


def bootstrap_windows():
    if os.name != "nt":
        return
    import ctypes
    from pathlib import Path
    # Some managed launchers omit variables used in registered OLE DB/ODBC DLL paths.
    # Existing values and the user's persistent/system environment are untouched.
    for name, identifier in (
        ("USERPROFILE", 40), ("APPDATA", 26), ("LOCALAPPDATA", 28),
        ("ProgramFiles", 38), ("ProgramFiles(x86)", 42),
        ("CommonProgramFiles", 43), ("CommonProgramFiles(x86)", 44),
    ):
        if not os.environ.get(name):
            buffer = ctypes.create_unicode_buffer(32768)
            if ctypes.windll.shell32.SHGetFolderPathW(None, identifier, None, 0, buffer) == 0:
                os.environ[name] = str(Path(buffer.value))
    buffer = ctypes.create_unicode_buffer(32768)
    if ctypes.windll.kernel32.GetWindowsDirectoryW(buffer, len(buffer)):
        for name in ("WINDIR", "SystemRoot"):
            if not os.environ.get(name):
                os.environ[name] = str(Path(buffer.value))
