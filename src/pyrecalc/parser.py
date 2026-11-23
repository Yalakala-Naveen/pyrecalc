"""A small recursive-descent parser for the Excel formula grammar.

Parses a formula string (without the leading ``=``) into a tuple-based AST:

    ("num", 1.5) ("str", "x") ("bool", True)
    ("ref", sheet_or_None, "A1") ("range", sheet_or_None, "A1", "B5")
    ("binop", op, left, right) ("unary", op, node) ("percent", node)
    ("call", "SUM", [arg, ...])

Cell references are normalized to upper-case A1 form without ``$``.
"""

from __future__ import annotations

import re

from .errors import FormulaSyntaxError

_TOKEN_RE = re.compile(
    r"""
    \s*(?:
      (?P<number>\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?
    | (?P<string>"(?:[^"]|"")*")
    | (?P<quoted_sheet>'(?:[^']|'')*')
    | (?P<cell>\$?[A-Za-z]{1,3}\$?\d+)
    | (?P<op><>|<=|>=|[+\-*/^&=<>(),:!%])
    | (?P<ident>[A-Za-z_\\][A-Za-z0-9_.]*)
    )
    """,
    re.VERBOSE,
)

_CELL_RE = re.compile(r"^\$?[A-Za-z]{1,3}\$?\d+$")


class _Tok:
    __slots__ = ("kind", "text")

    def __init__(self, kind: str, text: str) -> None:
        self.kind = kind
        self.text = text


def _tokenize(src: str) -> list[_Tok]:
    toks: list[_Tok] = []
    pos = 0
    while pos < len(src):
        m = _TOKEN_RE.match(src, pos)
        if not m or m.end() == pos:
            raise FormulaSyntaxError(f"cannot tokenize at {src[pos:pos + 10]!r}")
        pos = m.end()
        kind = m.lastgroup
        assert kind is not None
        toks.append(_Tok(kind, m.group(kind)))
    toks.append(_Tok("eof", ""))
    return toks


def normalize_ref(text: str) -> str:
    return text.replace("$", "").upper()


class _Parser:
    def __init__(self, toks: list[_Tok]) -> None:
        self.toks = toks
        self.i = 0

    def peek(self, ahead: int = 0) -> _Tok:
        return self.toks[min(self.i + ahead, len(self.toks) - 1)]

    def next(self) -> _Tok:
        tok = self.toks[self.i]
        self.i += 1
        return tok

    def expect_op(self, text: str) -> None:
        tok = self.next()
        if tok.kind != "op" or tok.text != text:
            raise FormulaSyntaxError(f"expected {text!r}, got {tok.text!r}")

    # grammar, lowest precedence first ---------------------------------
    def parse_expr(self):
        return self.parse_comparison()

    def parse_comparison(self):
        node = self.parse_concat()
        while self.peek().kind == "op" and self.peek().text in ("=", "<>", "<", "<=", ">", ">="):
            op = self.next().text
            node = ("binop", op, node, self.parse_concat())
        return node

    def parse_concat(self):
        node = self.parse_addsub()
        while self.peek().kind == "op" and self.peek().text == "&":
            self.next()
            node = ("binop", "&", node, self.parse_addsub())
        return node

    def parse_addsub(self):
        node = self.parse_muldiv()
        while self.peek().kind == "op" and self.peek().text in ("+", "-"):
            op = self.next().text
            node = ("binop", op, node, self.parse_muldiv())
        return node

    def parse_muldiv(self):
        node = self.parse_power()
        while self.peek().kind == "op" and self.peek().text in ("*", "/"):
            op = self.next().text
            node = ("binop", op, node, self.parse_power())
        return node

    def parse_power(self):
        node = self.parse_unary()
        while self.peek().kind == "op" and self.peek().text == "^":
            self.next()
            node = ("binop", "^", node, self.parse_unary())
        return node

    def parse_unary(self):
        tok = self.peek()
        if tok.kind == "op" and tok.text in ("+", "-"):
            self.next()
            return ("unary", tok.text, self.parse_unary())
        return self.parse_postfix()

    def parse_postfix(self):
        node = self.parse_primary()
        while self.peek().kind == "op" and self.peek().text == "%":
            self.next()
            node = ("percent", node)
        return node

    def parse_primary(self):
        tok = self.next()
        if tok.kind == "number":
            return ("num", float(tok.text))
        if tok.kind == "string":
            return ("str", tok.text[1:-1].replace('""', '"'))
        if tok.kind == "op" and tok.text == "(":
            node = self.parse_expr()
            self.expect_op(")")
            return node
        if tok.kind in ("ident", "quoted_sheet", "cell"):
            return self.parse_ident_like(tok)
        raise FormulaSyntaxError(f"unexpected token {tok.text!r}")

    def parse_ident_like(self, tok: _Tok):
        # sheet-qualified reference:  Sheet!A1  /  'My Sheet'!A1[:B5]
        if self.peek().kind == "op" and self.peek().text == "!":
            self.next()
            sheet = tok.text[1:-1].replace("''", "'") if tok.kind == "quoted_sheet" else tok.text
            ref_tok = self.next()
            if ref_tok.kind not in ("ident", "cell") or not _CELL_RE.match(ref_tok.text):
                raise FormulaSyntaxError(f"expected a cell reference after '!', got {ref_tok.text!r}")
            return self.maybe_range(sheet, normalize_ref(ref_tok.text))
        if tok.kind == "quoted_sheet":
            raise FormulaSyntaxError("quoted sheet name must be followed by '!'")
        upper = tok.text.upper()
        if upper in ("TRUE", "FALSE"):
            return ("bool", upper == "TRUE")
        # function call
        if self.peek().kind == "op" and self.peek().text == "(":
            self.next()
            args = []
            if not (self.peek().kind == "op" and self.peek().text == ")"):
                args.append(self.parse_expr())
                while self.peek().kind == "op" and self.peek().text == ",":
                    self.next()
                    args.append(self.parse_expr())
            self.expect_op(")")
            return ("call", upper, args)
        if _CELL_RE.match(tok.text):
            return self.maybe_range(None, normalize_ref(tok.text))
        # a bare name we do not know: keep as a name node -> #NAME? at eval
        return ("name", upper)

    def maybe_range(self, sheet, start: str):
        if self.peek().kind == "op" and self.peek().text == ":":
            self.next()
            end_tok = self.next()
            if end_tok.kind not in ("ident", "cell") or not _CELL_RE.match(end_tok.text):
                raise FormulaSyntaxError(f"expected a cell reference after ':', got {end_tok.text!r}")
            return ("range", sheet, start, normalize_ref(end_tok.text))
        return ("ref", sheet, start)


def parse(formula: str):
    """Parse a formula body (no leading '=') into an AST tuple."""
    parser = _Parser(_tokenize(formula.strip()))
    node = parser.parse_expr()
    if parser.peek().kind != "eof":
        raise FormulaSyntaxError(f"trailing input at {parser.peek().text!r}")
    return node
