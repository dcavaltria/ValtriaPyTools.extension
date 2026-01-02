# -*- coding: utf-8 -*-
"""Text helpers shared across Valtria PyTools (IronPython friendly)."""

try:
    unicode
except NameError:
    unicode = str  # type: ignore


def ensure_text(value):
    if value is None:
        return u""
    if isinstance(value, unicode):
        return value
    try:
        return unicode(value)
    except Exception:
        try:
            return str(value)
        except Exception:
            return u""


def safe_text(value):
    return ensure_text(value)


__all__ = ['ensure_text', 'safe_text']
