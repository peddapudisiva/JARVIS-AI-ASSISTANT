# Minimal version utilities to replace distutils.version for Python 3.12/3.13
# Implements LooseVersion/StrictVersion compatible enough for simple comparisons.
from packaging.version import Version

class _BaseVersion:
    def __init__(self, v):
        self.version = str(v)
        self._v = Version(self.version)
    def __repr__(self):
        return f"{self.__class__.__name__}('{self.version}')"
    def _coerce(self, other):
        if isinstance(other, _BaseVersion):
            return other._v
        return Version(str(other))
    def __eq__(self, other):
        return self._v == self._coerce(other)
    def __lt__(self, other):
        return self._v < self._coerce(other)
    def __le__(self, other):
        return self._v <= self._coerce(other)
    def __gt__(self, other):
        return self._v > self._coerce(other)
    def __ge__(self, other):
        return self._v >= self._coerce(other)

class LooseVersion(_BaseVersion):
    pass

class StrictVersion(_BaseVersion):
    pass
