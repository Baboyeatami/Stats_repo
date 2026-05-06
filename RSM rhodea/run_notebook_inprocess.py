from __future__ import annotations

import ast
import contextlib
import io
import os
import sys
import textwrap
import traceback
from pathlib import Path

os.environ.setdefault("MPLBACKEND", "Agg")

import nbformat as nbf


ROOT = Path(__file__).resolve().parent
NOTEBOOK = ROOT / "rhodea_garlic_peel_analysis.ipynb"


def as_text(value) -> str:
    if value is None:
        return ""
    if hasattr(value, "_repr_html_"):
        try:
            return str(value)
        except Exception:
            pass
    return repr(value)


def execute_cell(source: str, env: dict):
    source = textwrap.dedent(source).strip()
    stdout = io.StringIO()
    outputs = []

    def display(*values):
        for value in values:
            text = str(value)
            if hasattr(value, "to_string"):
                try:
                    text = value.to_string()
                except Exception:
                    text = str(value)
            print(text)

    env["display"] = display

    try:
        tree = ast.parse(source, mode="exec")
        last_expr = None
        if tree.body and isinstance(tree.body[-1], ast.Expr):
            last_expr = ast.Expression(tree.body[-1].value)
            tree.body = tree.body[:-1]
            ast.fix_missing_locations(tree)
            ast.fix_missing_locations(last_expr)

        with contextlib.redirect_stdout(stdout), contextlib.redirect_stderr(stdout):
            if tree.body:
                exec(compile(tree, filename=str(NOTEBOOK), mode="exec"), env)
            if last_expr is not None:
                result = eval(compile(last_expr, filename=str(NOTEBOOK), mode="eval"), env)
                if result is not None:
                    if hasattr(result, "to_string"):
                        try:
                            result_text = result.to_string()
                        except Exception:
                            result_text = as_text(result)
                    else:
                        result_text = as_text(result)
                    outputs.append(
                        nbf.v4.new_output(
                            output_type="execute_result",
                            data={"text/plain": result_text},
                            metadata={},
                            execution_count=None,
                        )
                    )

    except Exception:
        err = traceback.format_exc()
        outputs.append(
            nbf.v4.new_output(
                output_type="error",
                ename=sys.exc_info()[0].__name__,
                evalue=str(sys.exc_info()[1]),
                traceback=err.splitlines(),
            )
        )
        raise
    finally:
        text = stdout.getvalue()
        if text:
            outputs.insert(0, nbf.v4.new_output(output_type="stream", name="stdout", text=text))

    return outputs


def main() -> None:
    nb = nbf.read(NOTEBOOK, as_version=4)
    env = {"__name__": "__main__", "__file__": str(NOTEBOOK)}
    execution_count = 1
    for cell in nb.cells:
        if cell.cell_type != "code":
            continue
        cell.outputs = []
        cell.execution_count = execution_count
        outputs = execute_cell(cell.source, env)
        for output in outputs:
            if output.output_type == "execute_result":
                output.execution_count = execution_count
        cell.outputs = outputs
        execution_count += 1
    nbf.write(nb, NOTEBOOK)
    print(f"Executed {execution_count - 1} code cells in {NOTEBOOK}")


if __name__ == "__main__":
    main()
