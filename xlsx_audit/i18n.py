"""Tiny message catalogue (English default, Turkish opt-in).

The card ships with a Turkish CLI example, so findings can be rendered in
either language with ``--lang``.  Codes stay English everywhere: they are the
stable contract used by CI, JSON and SARIF.
"""

from __future__ import annotations

from typing import Callable, Dict

SUPPORTED = ("en", "tr")

_MESSAGES: Dict[str, Dict[str, str]] = {
    "en": {
        "broken_reference": "Broken formula reference found ({code}).",
        "error_value": "Cell holds the Excel error value {code}.",
        "formula_parse_error": "Formula could not be parsed statically.",
        "magic_number": "Hard-coded numeric constant {value} in formula.",
        "external_link": "External workbook reference: {target}",
        "external_link_cell": "Formula references an external workbook [{book}].",
        "external_url": "Formula calls out to a network resource ({target}).",
        "very_hidden_sheet": "Sheet is veryHidden: invisible in the Excel UI.",
        "hidden_sheet": "Sheet is hidden.",
        "metadata_present": "Document metadata is present ({keys}); review before publishing.",
        "circular_reference": "Circular reference detected: {path}",
        "omitted_cell": "Range {range} ends before the filled cell {cell} ({value}).",
        "hidden_defined_name": "Defined name {name!r} is hidden.",
    },
    "tr": {
        "broken_reference": "Kırık formül referansı bulundu ({code}).",
        "error_value": "Hücre Excel hata değeri içeriyor: {code}.",
        "formula_parse_error": "Formül statik olarak çözümlenemedi.",
        "magic_number": "Formül içine gömülü sabit sayı: {value}.",
        "external_link": "Dış çalışma kitabı bağlantısı: {target}",
        "external_link_cell": "Formül dış çalışma kitabına başvuruyor: [{book}].",
        "external_url": "Formül ağ kaynağına çağrı yapıyor ({target}).",
        "very_hidden_sheet": "Sayfa veryHidden olarak gizlenmiş: Excel arayüzünde görünmez.",
        "hidden_sheet": "Sayfa gizli.",
        "metadata_present": "Belge üst verisi mevcut ({keys}); yayınlamadan önce gözden geçirin.",
        "circular_reference": "Döngüsel başvuru tespit edildi: {path}",
        "omitted_cell": "{range} toplam aralığı dolu hücreyi kapsamıyor: {cell} ({value}).",
        "hidden_defined_name": "{name!r} tanımlı adı gizli.",
    },
}


def normalize_lang(lang: str | None) -> str:
    if not lang:
        return "en"
    lowered = lang.strip().lower().replace("_", "-")
    if lowered.startswith("tr"):
        return "tr"
    return "en"


def translator(lang: str | None) -> Callable[..., str]:
    """Return a ``t(key, **kwargs)`` callable for ``lang``."""
    code = normalize_lang(lang)
    catalogue = _MESSAGES[code]
    fallback = _MESSAGES["en"]

    def t(key: str, **kwargs: object) -> str:
        template = catalogue.get(key) or fallback.get(key) or key
        return template.format(**kwargs)

    return t
