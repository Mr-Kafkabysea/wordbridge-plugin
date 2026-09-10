"""Read selected text from an explicitly identified Word test document."""

import argparse
import json
from pathlib import Path

if __package__:
    from .check_active_document import check_active_document
else:
    from check_active_document import check_active_document


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--expected-document", required=True, type=Path)
    args = parser.parse_args()
    result = check_active_document(args.expected_document, include_selection=True)
    # Selected content is returned to the caller, never saved to a log file.
    print(json.dumps(result, ensure_ascii=True, indent=2))
    return 0 if result["status"] == "selected" else 1


if __name__ == "__main__":
    raise SystemExit(main())
