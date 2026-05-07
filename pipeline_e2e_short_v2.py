from __future__ import annotations

import sys

from pipeline_tc import main


if __name__ == "__main__":
    sys.argv.insert(1, "e2e_short_v2")
    raise SystemExit(main())
