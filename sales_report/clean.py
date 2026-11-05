"""Standardise product names across branches and flag values that look like typing errors."""
import difflib
import re
from collections import Counter

import pandas as pd

from . import config


def _key(text: str) -> str:
    return re.sub(r"[^a-z]", "", text.lower())


def standardise_products(df: pd.DataFrame) -> tuple[pd.DataFrame, list[tuple[str, str, int]], list[str]]:
    """Map every product spelling to the official name.

    Returns the data, the corrections made (raw name, official name, rows) and names that could not be matched.
    """
    official = {_key(p): p for p in config.PRODUCTS}
    aliases = {_key(a): p for a, p in config.PRODUCT_ALIASES.items()}
    cache, unknown = {}, []

    def match(raw: str) -> str:
        if raw in cache:
            return cache[raw]
        k = _key(raw)
        result = official.get(k) or aliases.get(k)
        if not result:
            close = difflib.get_close_matches(k, list(official), n=1, cutoff=0.8)
            result = official[close[0]] if close else None
        if not result:
            unknown.append(raw)
            result = raw.strip()
        cache[raw] = result
        return result

    df = df.copy()
    df["product"] = df["product_raw"].map(match)
    changed = df[df["product_raw"] != df["product"]]
    corrections = Counter(zip(changed["product_raw"], changed["product"], changed["branch"]))
    rows = [(raw, fixed, branch, n) for (raw, fixed, branch), n in sorted(corrections.items(), key=lambda kv: -kv[1])]
    return df, rows, sorted(set(unknown))


def flag_outliers(df: pd.DataFrame, threshold: float = 6.0) -> pd.DataFrame:
    """Daily quantities far outside what a branch normally sells of that product (robust z-score).

    Uses the median and MAD, which a single wrong value cannot distort. Flagged rows are kept,
    not deleted: a person should confirm them with the branch.
    """
    grouped = df.groupby(["branch", "product"])["qty"]
    median = grouped.transform("median")
    mad = grouped.transform(lambda s: (s - s.median()).abs().median()).replace(0, 1)
    z = (df["qty"] - median) / (1.4826 * mad)
    flagged = df[z.abs() > threshold].copy()
    flagged["typical"] = median[flagged.index].round(0).astype(int)
    flagged["z"] = z[flagged.index].round(1)
    return flagged
