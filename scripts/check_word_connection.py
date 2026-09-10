"""Read-only probe of an existing Word COM instance; never starts Word."""

import json

import pythoncom
import pywintypes
import win32com.client


MK_E_UNAVAILABLE = 0x800401E3


def check_word_connection() -> dict:
    """Return plain data, releasing COM references on the calling thread.

    An unavailable active object does not prove that Word is not running:
    registration, session boundaries or permissions can also prevent access.
    """
    initialized = False
    word = None
    stage = "initialize"
    try:
        pythoncom.CoInitializeEx(pythoncom.COINIT_APARTMENTTHREADED)
        initialized = True
        stage = "connect"
        word = win32com.client.GetActiveObject("Word.Application")
        stage = "read_version"
        version = str(word.Version)
        return {
            "status": "connected",
            "connected": True,
            "word_version": version,
            "message": "Connected to an existing Word instance; no documents accessed.",
        }
    except pywintypes.com_error as error:
        code = error.hresult & 0xFFFFFFFF
        unavailable = stage == "connect" and code == MK_E_UNAVAILABLE
        return {
            "status": "not_available" if unavailable else "com_error",
            "connected": False,
            "stage": stage,
            "hresult": f"0x{code:08X}",
            "message": (
                "No accessible active Word instance. Word may be closed or not "
                "registered for this session. Open Word, switch to another app, "
                "then try again."
                if unavailable
                else "Word COM probe failed. Check Word availability and session permissions."
            ),
        }
    finally:
        # Drop the application reference before uninitializing this thread.
        word = None
        if initialized:
            pythoncom.CoUninitialize()


def main() -> int:
    result = check_word_connection()
    print(json.dumps(result, ensure_ascii=True, indent=2))
    return 0 if result["connected"] else 1


if __name__ == "__main__":
    raise SystemExit(main())
