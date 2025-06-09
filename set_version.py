#!/usr/bin/env python3
"""Updates package version number, and commits and tags repository."""

import argparse
import re
from pathlib import Path
from subprocess import PIPE, STDOUT, run  # noqa: S404

import pendulum
from semver import compare

PKG_DIR = Path(__file__).parent
PKG_NAME = next(
    _f for _f in (PKG_DIR / "src").iterdir() if not _f.name.startswith(".")
).name

# Set up the argument parser
parser = argparse.ArgumentParser(
    description="Updates package version number, and commits and tags repository. User must specify `full` or `patch` level update."
)
parser.add_argument(
    "update_level",
    type=str,
    choices=["full", "patch"],
    help="Whether `full` or `patch` level version update.",
)

args = parser.parse_args()

tsn = pendulum.today()

rc = run(  # noqa: S603
    ["/usr/bin/git", "status", "-uno"],
    check=True,
    stdout=PIPE,
    stderr=STDOUT,
    text=True,
)
if "nothing to commit" not in rc.stdout:
    raise RuntimeError(
        "Repository has uncommitted changes. Commit changes before updating package version."
    )

# Update README.rst from the docs version
# strip_sphinx_pat = re.compile(r":(attr|class|func|meth|mod):")
# (PKG_DIR / "README.rst").write_text(
#     strip_sphinx_pat.sub(
#         r":code:", (PKG_DIR / "docs" / "source" / "README.rst").read_text()
#     )
# )
# (PKG_DIR / "README.rst").write_text(
#     re.sub(
#         r":ref:`(?P<ref>.*?) (?:<.*?>)`",
#         r"\g<ref>",
#         (PKG_DIR / "README.rst").read_text(),
#     )

# re.sub(r":ref:`(?P<ref>.*) (?:<.*>)`", r"\g<ref>", "See, for example, :ref:`Coate (2011) <coate2011>`.")

# Update license
license_path = PKG_DIR / "docs" / "source" / "license.rst"
license_path.write_text(
    re.sub(
        r"Copyright (?P<byr>\d{4})-\d{4} (?P<name>S\. Murthy Kambhampaty)",
        rf"Copyright \g<byr>-{tsn.year} \g<name>",
        license_path.read_text(),
    )
)


def _get_pkg_version() -> str:
    return run(  # noqa: S603
        ["./.venv/bin/poetry", "version", "-s"], stdout=PIPE, text=True, check=True
    ).stdout.strip()


pkg_ver = _get_pkg_version()
sem_ver = f"{tsn.year}.{tsn.toordinal()}.0"

# Update pyproject.toml
match args.update_level:
    case "patch":
        run(["./.venv/bin/poetry", "version", "patch"], check=True)  # noqa: S603
        sem_ver = _get_pkg_version()
    case "full":
        if compare(sem_ver, pkg_ver) <= 0:
            raise ValueError(
                f"Package version, {pkg_ver} at or above version, {sem_ver}. Perhaps update patch-level."
            )

        run(["./.venv/bin/poetry", "version", sem_ver], check=True)  # noqa: S603

# Update package's main __init__.py
pkg_init_path = PKG_DIR / "src" / PKG_NAME / "__init__.py"
pkg_init_path.write_text(
    re.sub(
        rf'(?m)^VERSION = "{pkg_ver}"$',
        f'VERSION = "{sem_ver}"',
        pkg_init_path.read_text(),
    )
)

# Commit, tag and push
run(  # noqa: S603
    [
        "/usr/bin/git",
        "commit",
        "-i",
        "README.rst",
        "pyproject.toml",
        f"{pkg_init_path}",
        "docs/source/license.rst",
        "-m",
        f'"chore({tsn.to_date_string()}): update version"',
    ],
    check=True,
)
run(["/usr/bin/git", "tag", f"{sem_ver}"], check=True)  # noqa: S603
run(["/usr/bin/git", "push"], check=True)  # noqa: S603
run(["/usr/bin/git", "push", "--tags"], check=True)  # noqa: S603
