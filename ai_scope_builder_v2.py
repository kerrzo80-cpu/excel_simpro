#!/usr/bin/env python3
"""
AI Scope Builder v2.

Fixes object detection from v1:
- Stops detecting "pan" inside "panel".
- Uses whole-word matching for single-word object aliases.
- Still uses phrase matching for multi-word terms like "outside tap".

Usage:
  python3 ai_scope_builder_v2.py "replace radiator with new double panel radiator"
"""

from __future__ import annotations

import re
import ai_scope_builder as base


def infer_objects_safe(description: str) -> set[str]:
    text = base.norm(description)
    token_set = base.words(description)
    objects: set[str] = set()

    aliases = {
        "basin": ["basin", "sink", "whb", "wash hand basin", "vanity"],
        "radiator": ["radiator", "rad", "rads", "trv"],
        "toilet": ["toilet", "wc", "loo", "pan", "cistern", "syphon", "siphon"],
        "outside tap": ["outside tap", "garden tap", "external tap"],
        "pipework": ["pipe", "pipes", "pipework", "supplies", "hot", "cold"],
        "shower": ["shower"],
        "bath": ["bath"],
        "tap": ["tap", "taps", "mixer"],
        "waste": ["waste", "trap"],
    }

    for obj, terms in aliases.items():
        for term in terms:
            n_term = base.norm(term)
            if " " in n_term:
                if re.search(rf"\b{re.escape(n_term)}\b", text):
                    objects.add(obj)
                    break
            else:
                if n_term in token_set:
                    objects.add(obj)
                    break

    return objects


# Patch the v1 module so all existing functions use the safer detector.
base.infer_objects = infer_objects_safe


if __name__ == "__main__":
    raise SystemExit(base.main())
