"""Retarget worksheet references without changing quoted text inside preserved formulas."""

import re

from openpyxl.formula.tokenizer import Tokenizer


def rebind_formula(formula: str, routes: dict[str, str]) -> str:
    parsed = Tokenizer(formula)
    for token in parsed.items:
        if token.type != "OPERAND" or token.subtype != "RANGE":
            continue
        for source, target in routes.items():
            quoted_source = source.replace("'", "''")
            quoted_target = target.replace("'", "''")
            token.value = token.value.replace(f"'{quoted_source}'!", f"'{quoted_target}'!")
            token.value = re.sub(
                rf"(?<![\w']){re.escape(source)}!", f"'{quoted_target}'!", token.value
            )
    return "=" + "".join(token.value for token in parsed.items)
