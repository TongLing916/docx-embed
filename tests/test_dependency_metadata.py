from __future__ import annotations

import tomllib
from pathlib import Path


def test_markitdown_runtime_dependency_requests_only_docx_extra() -> None:
    project = tomllib.loads(Path("pyproject.toml").read_text(encoding="utf-8"))[
        "project"
    ]

    assert "markitdown[all]>=0.1.6" not in project["dependencies"]
    assert "markitdown[docx]>=0.1.6" in project["dependencies"]


def test_console_script_and_compat_examples_are_packaged() -> None:
    config = tomllib.loads(Path("pyproject.toml").read_text(encoding="utf-8"))

    assert config["project"]["scripts"]["edp"] == "edp.cli:main"
    assert "examples*" in config["tool"]["setuptools"]["packages"]["find"]["include"]
