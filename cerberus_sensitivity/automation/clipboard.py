from __future__ import annotations

import csv
import io
from typing import List, Dict


def parse_tabulated_tsv(raw_text: str) -> List[Dict[str, str]]:
    if not raw_text:
        return []
    reader = csv.DictReader(io.StringIO(raw_text), delimiter="\t")
    return [dict(row) for row in reader]
