"""The LaTeX commands docx4j's math converters share. CR-011 section 6.

docx4j's ``org.docx4j.model.math.LatexSymbols`` at ``198d015a3`` (its CR-046
D3), verbatim: the n-ary operators and the symbols by command, and the
function names LaTeX sets upright.
"""

from __future__ import annotations

__all__ = ["FUNCTION_NAMES", "NARY", "SYMBOLS"]

#: The function names LaTeX sets upright (``\\sin``, ``\\log``, ...).
FUNCTION_NAMES = frozenset(
    {
        "sin", "cos", "tan", "cot", "sec", "csc",
        "arcsin", "arccos", "arctan", "sinh", "cosh", "tanh",
        "log", "ln", "lg", "exp", "max", "min", "sup", "inf",
        "lim", "arg", "det", "dim", "gcd", "mod",
    }
)  # fmt: skip

#: The n-ary operators by command (``sum`` to U+2211, ...).
NARY: dict[str, str] = {
    "sum": "∑",
    "prod": "∏",
    "coprod": "∐",
    "int": "∫",
    "iint": "∬",
    "iiint": "∭",
    "oint": "∮",
    "bigcup": "⋃",
    "bigcap": "⋂",
}

_GREEK = (
    ("alpha", "α"), ("beta", "β"), ("gamma", "γ"), ("delta", "δ"),
    ("epsilon", "ϵ"), ("varepsilon", "ε"), ("zeta", "ζ"), ("eta", "η"),
    ("theta", "θ"), ("vartheta", "ϑ"), ("iota", "ι"), ("kappa", "κ"),
    ("lambda", "λ"), ("mu", "μ"), ("nu", "ν"), ("xi", "ξ"),
    ("pi", "π"), ("varpi", "ϖ"), ("rho", "ρ"), ("varrho", "ϱ"),
    ("sigma", "σ"), ("varsigma", "ς"), ("tau", "τ"), ("upsilon", "υ"),
    ("phi", "ϕ"), ("varphi", "φ"), ("chi", "χ"), ("psi", "ψ"),
    ("omega", "ω"),
    ("Gamma", "Γ"), ("Delta", "Δ"), ("Theta", "Θ"), ("Lambda", "Λ"),
    ("Xi", "Ξ"), ("Pi", "Π"), ("Sigma", "Σ"), ("Upsilon", "Υ"),
    ("Phi", "Φ"), ("Psi", "Ψ"), ("Omega", "Ω"),
)  # fmt: skip

_TABLE = (
    ("pm", "±"), ("mp", "∓"), ("times", "×"), ("cdot", "⋅"), ("div", "÷"),
    ("ast", "∗"), ("star", "⋆"), ("circ", "∘"), ("bullet", "•"),
    ("oplus", "⊕"), ("ominus", "⊖"), ("otimes", "⊗"), ("oslash", "⊘"),
    ("le", "≤"), ("leq", "≤"), ("ge", "≥"), ("geq", "≥"),
    ("ne", "≠"), ("neq", "≠"), ("approx", "≈"), ("sim", "∼"),
    ("simeq", "≃"), ("cong", "≅"), ("equiv", "≡"), ("propto", "∝"),
    ("ll", "≪"), ("gg", "≫"), ("prec", "≺"), ("succ", "≻"),
    ("to", "→"), ("rightarrow", "→"), ("leftarrow", "←"),
    ("Rightarrow", "⇒"), ("Leftarrow", "⇐"),
    ("leftrightarrow", "↔"), ("Leftrightarrow", "⇔"),
    ("mapsto", "↦"), ("uparrow", "↑"), ("downarrow", "↓"),
    ("longrightarrow", "⟶"), ("longleftarrow", "⟵"),
    ("longleftrightarrow", "⟷"), ("Longrightarrow", "⟹"),
    ("Longleftarrow", "⟸"), ("Longleftrightarrow", "⟺"),
    ("implies", "⟹"), ("iff", "⟺"), ("mid", "∣"),
    ("infty", "∞"), ("partial", "∂"), ("nabla", "∇"),
    ("hbar", "ℏ"), ("ell", "ℓ"), ("Re", "ℜ"), ("Im", "ℑ"),
    ("aleph", "ℵ"), ("wp", "℘"), ("prime", "′"), ("degree", "°"),
    ("cdots", "⋯"), ("ldots", "…"), ("dots", "…"), ("vdots", "⋮"),
    ("ddots", "⋱"),
    ("in", "∈"), ("notin", "∉"), ("ni", "∋"),
    ("subset", "⊂"), ("subseteq", "⊆"), ("supset", "⊃"), ("supseteq", "⊇"),
    ("cup", "∪"), ("cap", "∩"), ("setminus", "∖"),
    ("emptyset", "∅"), ("varnothing", "∅"),
    ("forall", "∀"), ("exists", "∃"), ("nexists", "∄"), ("neg", "¬"),
    ("wedge", "∧"), ("land", "∧"), ("vee", "∨"), ("lor", "∨"),
    ("angle", "∠"), ("perp", "⊥"), ("parallel", "∥"), ("mid", "∣"),
    ("langle", "⟨"), ("rangle", "⟩"),
)  # fmt: skip

#: The symbols by command: Greek letters, operators, relations, arrows (``alpha`` to U+03B1, ...).
SYMBOLS: dict[str, str] = dict(_GREEK + _TABLE)
