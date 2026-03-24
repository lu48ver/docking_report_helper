import os
import sys


def get_app_base_dir() -> str:
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def get_app_path(*parts: str) -> str:
    return os.path.join(get_app_base_dir(), *parts)
