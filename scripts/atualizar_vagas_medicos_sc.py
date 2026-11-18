#!/usr/bin/env python3
"""Collect public SC medical openings and refresh the site's JSON feed."""

from __future__ import annotations

import html
import json
import re
import sys
from datetime import date, datetime
from pathlib import Path
from urllib.error import HTTPError, URLError
from urllib.parse import urljoin
from urllib.request import Request, urlopen
from zoneinfo import ZoneInfo


ROOT = Path(__file__).resolve().parents[1]
OUTPUT = ROOT / "vagas-medicos-sc.json"
TZ = ZoneInfo("America/Sao_Paulo")
TODAY = datetime.now(TZ).date()

SOURCES = [
    {
        "name": "SES/SC",
        "institution": "Secretaria de Estado da Saúde de Santa Catarina",
        "url": "https://www.saude.sc.gov.br/index.php/pt/servicos/concursos-e-selecoes/processos-seletivos-ses",
    },
    {
        "name": "Diário Oficial dos Municípios de SC",
        "institution": "Município de Santa Catarina",
        "url": "https://static.dom.sc.gov.br/?q=m%C3%A9dico&r=pesquisa%2Fsite",
    },
    {
        "name": "SIMESC",
        "institution": "Sindicato dos Médicos do Estado de Santa Catarina",
        "url": "https://www.simesc.org.br/empregos/default.aspx",
    },
    {
        "name": "Hospital Santa Isabel",
        "institution": "Hospital Santa Isabel / Rede Santa Catarina",
        "url": "https://redesantacatarina.org.br/hospital/santaisabel/SitePages/institucional/trabalhe-conosco.aspx",
    },
]

VACANCY = re.compile(
    r"vaga|vagas|plant[aã]o|contrata|contratando|oportunidade|m[eé]dic[oa]|"
    r"cl[ií]nic[oa]|pronto atendimento|urg[eê]ncia|escala|credenciamento|processo seletivo",
    re.I,
)
MEDICINE = re.compile(r"m[eé]dic[oa]|medicina", re.I)
SANTA_CATARINA = re.compile(
    r"santa catarina|\bsc\b|florian[oó]polis|s[aã]o jos[eé]|palho[cç]a|bigua[cç]u|"
    r"itaj[aá]i|balne[aá]rio cambori[uú]|navegantes|brusque|blumenau|joinville|"
    r"jaragu[aá] do sul|chapec[oó]|conc[oó]rdia|crici[uú]ma|tubar[aã]o|lages|"
    r"rio do sul|joa[cç]aba|maf[r]?a|ca[cç]ador|s[aã]o bento do sul|ararangu[aá]|"
    r"itapema|indaial|garopaba|grande florian[oó]polis",
    re.I,
)
DATE = re.compile(r"\b(\d{1,2})/(\d{1,2})(?:/(20\d{2}|\d{2}))?\b")
ANCHOR = re.compile(r"<a\b[^>]*href\s*=\s*([\"'])(.*?)\1[^>]*>([\s\S]*?)</a\s*>", re.I)
TAG = re.compile(r"<[^>]+>")

CITY_NAMES = [
    "Florianópolis", "São José", "Palhoça", "Biguaçu", "Itajaí", "Balneário Camboriú",
    "Navegantes", "Brusque", "Blumenau", "Joinville", "Jaraguá do Sul", "Chapecó",
    "Concórdia", "Criciúma", "Tubarão", "Lages", "Rio do Sul", "Joaçaba", "Mafra",
    "Caçador", "São Bento do Sul", "Araranguá", "Itapema", "Indaial", "Garopaba",
]
SPECIALTIES = [
    "cardiologia pediátrica", "medicina do trabalho", "neurologia pediátrica",
    "reumatologia pediátrica", "psiquiatria", "pediatria", "clínica médica",
    "clínico geral", "medicina de emergência", "médico de voo", "aeromédico",
    "medicina intensiva", "ginecologia", "obstetrícia", "cirurgia geral",
]


def clean_text(value: str) -> str:
    return re.sub(r"\s+", " ", html.unescape(TAG.sub(" ", value))).strip()


def fetch(url: str) -> str:
    request = Request(url, headers={"User-Agent": "HandoffVagasMedicasSC/1.0 (+https://github.com/yarisdd/handoff-vagas)"})
    with urlopen(request, timeout=25) as response:
        if response.status >= 400:
            raise RuntimeError(f"HTTP {response.status}")
        return response.read().decode(response.headers.get_content_charset() or "utf-8", errors="replace")


def parse_deadline(text: str) -> date | None:
    start = re.search(r"inscri[cç][oõ]es?|candidaturas?|prazo", text, re.I)
    if not start:
        return None
    window = text[start.start() : start.start() + 260]
    dates = list(DATE.finditer(window))
    if not dates:
        return None
    # A listing usually states the opening date first and the closing date last.
    match = dates[-1]
    closing_phrase = window[max(0, match.start() - 36) : match.end()]
    date_range = False
    if len(dates) >= 2:
        between_dates = window[dates[-2].end() : match.start()]
        date_range = bool(re.fullmatch(r"\s*(?:a|até|ate|[-–—])\s*", between_dates, re.I))
    if not date_range and not re.search(r"\bat[eé]\b|encerr|t[eé]rmino|final", closing_phrase, re.I):
        return None
    day, month = int(match.group(1)), int(match.group(2))
    year_text = match.group(3)
    year = int(year_text) if year_text else TODAY.year
    if year_text and len(year_text) == 2:
        year += 2000
    try:
        return date(year, month, day)
    except ValueError:
        return None


def city_from(text: str) -> str:
    for city in CITY_NAMES:
        if re.search(re.escape(city), text, re.I):
            return city
    if re.search(r"grande florian[oó]polis", text, re.I):
        return "Grande Florianópolis"
    return ""


def specialty_from(text: str) -> str:
    for specialty in SPECIALTIES:
        if re.search(re.escape(specialty), text, re.I):
            return specialty.title()
    return "Medicina"


def title_from(anchor_title: str, context: str, specialty: str, city: str) -> str:
    title = anchor_title.strip()
    if title and len(title) > 5 and not re.fullmatch(r"(edital|saiba mais|acesse|clique aqui|download)", title, re.I):
        return title[:140]
    pss = re.search(r"(?:PSS|edital|processo seletivo)\s*(?:n[ºo.]?\s*)?[0-9./-]+", context, re.I)
    prefix = pss.group(0).strip() if pss else "Vaga médica"
    suffix = f" — {specialty}"
    if city:
        suffix += f" · {city}"
    return (prefix + suffix)[:140]


def collect_source(source: dict[str, str]) -> tuple[list[dict], bool]:
    page = source["url"]
    raw = fetch(page)
    found: dict[str, dict] = {}
    for match in ANCHOR.finditer(raw):
        link = urljoin(page, html.unescape(match.group(2).strip()))
        if link.rstrip("/") == page.rstrip("/") or not link.startswith(("http://", "https://")):
            continue
        if re.search(r"\.(?:png|jpe?g|gif|svg|css|js)(?:[?#]|$)", link, re.I):
            continue
        start = max(0, match.start() - 900)
        end = min(len(raw), match.end() + 900)
        context = clean_text(raw[start:end])
        anchor_title = clean_text(match.group(3))
        if anchor_title and anchor_title not in context:
            context = f"{anchor_title} {context}"
        if not (VACANCY.search(context) and MEDICINE.search(context) and SANTA_CATARINA.search(context)):
            continue
        deadline = parse_deadline(context)
        if not deadline or deadline < TODAY:
            continue
        city = city_from(context)
        specialty = specialty_from(context)
        summary_start = MEDICINE.search(context).start()
        summary = context[max(0, summary_start - 100) : summary_start + 420]
        found[link] = {
            "id_origem": link,
            "fonte_automatica": source["name"],
            "titulo": title_from(anchor_title, context, specialty, city),
            "especialidade": specialty,
            "instituicao": source["institution"],
            "cidade": city,
            "uf": "SC",
            "modalidade": "Conforme edital",
            "resumo": summary[:420],
            "link": link,
            "prazo": deadline.strftime("%d/%m/%Y"),
            "data_fim": deadline.isoformat(),
            "atualizado_em": TODAY.isoformat(),
        }
    return list(found.values()), True


def read_existing() -> list[dict]:
    if not OUTPUT.exists():
        return []
    value = json.loads(OUTPUT.read_text(encoding="utf-8") or "[]")
    if not isinstance(value, list):
        raise ValueError("vagas-medicos-sc.json must contain a JSON array")
    return value


def main() -> int:
    candidates: dict[str, dict] = {}
    successful = 0
    errors: list[str] = []
    for source in SOURCES:
        try:
            items, _ = collect_source(source)
            successful += 1
            candidates.update({item["id_origem"]: item for item in items})
            print(f"{source['name']}: {len(items)} oportunidades válidas")
        except (HTTPError, URLError, TimeoutError, UnicodeError, RuntimeError) as exc:
            errors.append(f"{source['name']}: {exc}")
            print(f"{source['name']}: erro de coleta ({exc})", file=sys.stderr)

    if successful == 0:
        print("Todas as fontes falharam; feed anterior preservado.", file=sys.stderr)
        return 1

    by_link: dict[str, dict] = {}
    for job in read_existing():
        link = str(job.get("id_origem") or job.get("link") or "").strip()
        if not link:
            continue
        end = job.get("data_fim")
        if end:
            if str(end) >= TODAY.isoformat():
                by_link[link] = job
        elif not job.get("fonte_automatica"):
            # Preserve undated manual entries until they are reviewed.
            by_link[link] = job
    by_link.update(candidates)
    result = sorted(by_link.values(), key=lambda item: item.get("data_fim", "9999-12-31"))
    OUTPUT.write_text(json.dumps(result, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(f"Feed atualizado: {len(result)} oportunidades válidas.")
    if errors:
        print(f"Fontes com falha: {len(errors)}; oportunidades existentes válidas foram preservadas.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
