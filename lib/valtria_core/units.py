# -*- coding: utf-8 -*-
"""Unit conversion helpers shared across Valtria PyTools."""

_MM_PER_FOOT = 304.8


def feet_to_mm(value):
    try:
        return float(value) * _MM_PER_FOOT
    except Exception:
        return 0.0


__all__ = ['feet_to_mm']
