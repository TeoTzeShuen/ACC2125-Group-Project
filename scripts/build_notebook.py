"""Convert scripts/notebook_source.txt into Code_GroupXX.ipynb and execute it in place."""
import sys
from pathlib import Path
import nbformat
from nbclient import NotebookClient

ROOT = Path(__file__).resolve().parent.parent
src = (ROOT / "scripts" / "notebook_source.txt").read_text(encoding="utf-8")

cells = []
for block in src.split("# %%")[1:]:
    header, _, body = block.partition("\n")
    body = body.strip("\n")
    if header.strip() == "[markdown]":
        lines = [l[2:] if l.startswith("# ") else ("" if l == "#" else l) for l in body.splitlines()]
        cells.append(nbformat.v4.new_markdown_cell("\n".join(lines)))
    else:
        cells.append(nbformat.v4.new_code_cell(body))

nb = nbformat.v4.new_notebook(cells=cells)
nb.metadata["kernelspec"] = {"name": "python3", "display_name": "Python 3", "language": "python"}
out = ROOT / "Code_GroupXX.ipynb"
if "--no-run" not in sys.argv:
    NotebookClient(nb, timeout=1800, kernel_name="python3", resources={"metadata": {"path": str(ROOT)}}).execute()
nbformat.write(nb, out)
print("Wrote", out)
