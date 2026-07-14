"""XLSX workbook parsing and asset extraction.

This subpackage provides:

* :func:`parse_xlsx_package` — parse an XLSX workbook into structured tables
  with cell metadata, sub-table splitting, and asset annotations.
* :func:`extract_xlsx_assets` — extract images, charts, OLE objects, and
  equations from the XLSX drawing relationship tree.
"""

from edp.xlsx.parser import parse_xlsx_package
from edp.xlsx.assets import XlsxAssetCollection, extract_xlsx_assets

__all__ = [
    "XlsxAssetCollection",
    "extract_xlsx_assets",
    "parse_xlsx_package",
]
