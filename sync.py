#!/usr/bin/env python3
"""Entry point: python sync.py [--apply] [--days N] [--check]

Dry-run by default. See README.md.
"""

from __future__ import annotations

import sys

from outlook_xurrent.cli import main

if __name__ == "__main__":
    sys.exit(main())
