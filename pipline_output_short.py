from __future__ import annotations

import sys

from pipeline_tc import main


if __name__ == "__main__":
    sys.argv.insert(1, "output_short")
    raise SystemExit(main())
