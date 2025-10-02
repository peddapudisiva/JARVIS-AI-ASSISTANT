# Minimal distutils shim for Python 3.12/3.13 compatibility
# Provides only what some packages import (distutils.version.LooseVersion)
from .version import LooseVersion, StrictVersion  # noqa: F401
