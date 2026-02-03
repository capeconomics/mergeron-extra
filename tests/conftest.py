from datetime import datetime
from typing import Any

import numpy as np
from numpy.typing import NDArray

np.set_printoptions(precision=18)


def timestamper() -> str:
    return f"{datetime.now().strftime("%F %T.%f")} |>  "

