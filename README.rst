mergeron-extra: Support package for users of `mergeron`
=======================================================

Modules potentially useful to users of the :code:`mergeron` package. 

The module :code:`mergeron_extra.proportions_tests` provides methods for estimating confidence intervals for proportions and for contrasts (differences) in proportions. This module is coded for conformance to the literature and to results from the corresponding modules in :code:`R`. Although written from scratch, the APIs implemented in the module included here are designed for consistency with the APIs in, :code:`statsmodels.stats.proportion` from the package, :code:`statsmodels` (https://pypi.org/project/statsmodels/). To access these directly:

.. code-block:: python

    import mergeron_extra.proportions_tests as prci

Module :code:`mergeron_extra.xlsxw_helper` is useful for writing highly formatted output to spreadsheets with xlsx format. (Does not re-export the parent package, `xlsxwriter`.) The class, :code:`mergeron_extra.xlsxw_helper.CFmt` and function, :code:`mergeron_extra.xlsxw_helper.array_to_sheet` are of particular interest, and can be accessed as :code:`xlh.CFmt` and :code:`xlh.array_to_sheet` with the following import:

.. code-block:: python

    import mergeron_extra.xlsxw_helper as xlh

A recent version of Paul Tol's python module, :code:`tol_colors.py`, which provides high-contrast color schemes for making displays with improved visibility for individuals with color-blindness, is redistributed within this package. Other than re-formatting and type annotation, the :code:`mergeron_extra.tol_colors` module is re-distributed as downloaded from, https://personal.sron.nl/~pault/data/tol_colors.py. The :code:`tol_colors.py` module is distributed under the Standard 3-clause BSD license. To access the :code:`mergeron_extra.tol_colors` module directly:

.. code-block:: python

    import mergeron_extra.tol_colors as ptc

.. image:: https://img.shields.io/endpoint?url=https://python-poetry.org/badge/v0.json
   :alt: Poetry
   :target: https://python-poetry.org/

.. image:: https://img.shields.io/endpoint?url=https://raw.githubusercontent.com/astral-sh/ruff/main/assets/badge/v2.json
   :alt: Ruff
   :target: https://github.com/astral-sh/ruff

.. image:: https://www.mypy-lang.org/static/mypy_badge.svg
   :alt: Checked with mypy
   :target: https://mypy-lang.org/

.. image:: https://img.shields.io/badge/License-MIT-yellow.svg
   :alt: License: MIT
   :target: https://opensource.org/licenses/MIT

