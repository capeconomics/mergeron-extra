from __future__ import annotations

import numpy as np
from numpy.typing import NDArray

VERSION = "2026.739650.0"

__version__ = VERSION

np.set_printoptions(precision=24, floatmode="fixed")

type ArrayBoolean = NDArray[np.bool_]
type ArrayFloat = NDArray[np.half | np.single | np.double]
type ArrayINT = NDArray[np.intp]

type ArrayDouble = NDArray[np.double]
type ArrayBIGINT = NDArray[np.int64]
