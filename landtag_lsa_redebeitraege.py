#!/usr/bin/env python3
"""
Extrahiert Redebeitraege aus den Plenarprotokollen (Stenografische Berichte) des
Landtags von Sachsen-Anhalt fuer mehrere Wahlperioden und speichert sie als Excel-Datei.

Datenquelle: PDF-Plenarprotokolle unter landtag.sachsen-anhalt.de / padoka.landtag.sachsen-anhalt.de
Beispiel-URLs (bestaetigt):
  https://landtag.sachsen-anhalt.de/fileadmin/files/plenum/wp6/042stzg.pdf
  https://landtag.sachsen-anhalt.de/fileadmin/files/plenum/wp7/105stzg.pdf
  https://padoka.landtag.sachsen-anhalt.de/files/plenum/wp7/085stzg.pdf

Verwendung:
  pip install requests pdfplumber openpyxl
  python landtag_lsa_redebeitraege.py --test      # schneller Funktionscheck (offline + 1 PDF)
  python landtag_lsa_redebeitraege.py             # voller Lauf ueber WP 6, 7, 8
"""
from __future__ import annotations

import argparse
import json
import os
import re
import sys
import time
from dataclasses import asdict, dataclass, field
from pathlib import Path
from typing import Iterator, Optional

import requests

try:
    import pdfplumber
except ImportError:
    pdfplumber = None

try:
    from openpyxl import Workbook
    from openpyxl.utils import get_column_letter
except ImportError:
    Workbook = None


USER_AGENT = "landtag-lsa-redebeitraege-script/1.0 (Recherche-Tool, Kontakt: siehe Repo-Owner)"

# Mehrere bekannte Host/Pfad-Varianten; werden der Reihe nach probiert.
PDF_URL_TEMPLATES = [
    "https://www.landtag.sachsen-anhalt.de/fileadmin/files/plenum/wp{wp}/{nr:03d}stzg.pdf",
    "https://landtag.sachsen-anhalt.de/fileadmin/files/plenum/wp{wp}/{nr:03d}stzg.pdf",
    "https://padoka.landtag.sachsen-anhalt.de/files/plenum/wp{wp}/{nr:03d}stzg.pdf",
]

MAX_SESSIONS_PER_WP = 250
MAX_CONSECUTIVE_MISSES = 8

DATE_RE = re.compile(r"\b(\d{1,2}\.\d{1,2}\.\d{4})\b")

# Ganze Zeilen, die nur aus einer Regieanweisung/einem Zwischenruf bestehen,
# z. B. "(Beifall bei der CDU)" oder "(Zuruf von der SPD: Genau!)".
STAGE_DIRECTION_RE = re.compile(r"^\(.*\)$")

# Erkennt eine Redner-Kopfzeile am Absatzanfang, z. B.:
#   "Guido Kosmehl (FDP):"
#   "Praesidentin Gabriele Brakebusch:"
#   "Petra Grimm-Benne (Ministerin fuer Arbeit, Soziales und Integration):"
SPEAKER_RE = re.compile(
    r"^(?P<name>"
    r"(?:Präsidentin|Präsident|Vizepräsidentin|Vizepräsident|Alterspräsident|Alterspräsidentin)?\s*"
    r"[A-ZÄÖÜ][\wÄÖÜäöüß.\-]*(?:\s+(?:von|van|de|der)?\s*[A-ZÄÖÜ][\wÄÖÜäöüß.\-]*){0,4}"
    r")"
    r"\s*(?:\((?P<rolle>[^()]{1,120})\))?"
    r"\s*:\s*(?P<rest>.*)$"
)


@dataclass
class Redebeitrag:
    wahlperiode: int
    sitzung: int
    datum: str
    nr: int
    redner: str
    rolle_fraktion: str
    text: str
    quelle_url: str


def build_session() -> requests.Session:
    s = requests.Session()
    s.headers.update({"User-Agent": USER_AGENT})
    return s


def get_with_retries(session: requests.Session, url: str, tries: int = 3, timeout: float = 30) -> Optional[requests.Response]:
    """Holt eine URL. Verbindungsfehler (Netzwerksperre, DNS, ...) werden NICHT
    wiederholt, da sie sofort und wiederholt gleich fehlschlagen wuerden.
    Nur transiente Serverfehler (5xx/429) werden mit kurzer Pause erneut versucht."""
    for attempt in range(tries):
        try:
            resp = session.get(url, timeout=timeout)
        except requests.RequestException as exc:
            print(f"  [WARN] {url} -> {exc}", file=sys.stderr)
            return None
        if resp.status_code in (429, 500, 502, 503, 504) and attempt < tries - 1:
            time.sleep(2 * (attempt + 1))
            continue
        return resp
    return None


def download_protocol(session: requests.Session, wp: int, nr: int, cache_dir: Path) -> Optional[tuple[Path, str]]:
    """Laedt ein Plenarprotokoll-PDF herunter (mit lokalem Cache). Gibt (Pfad, Quelle-URL) oder None zurueck."""
    cache_dir.mkdir(parents=True, exist_ok=True)
    local_path = cache_dir / f"wp{wp}_{nr:03d}stzg.pdf"
    if local_path.exists() and local_path.stat().st_size > 0:
        return local_path, f"(cache) {local_path.name}"

    for template in PDF_URL_TEMPLATES:
        url = template.format(wp=wp, nr=nr)
        resp = get_with_retries(session, url)
        if resp is not None and resp.status_code == 200 and resp.content[:4] == b"%PDF":
            local_path.write_bytes(resp.content)
            return local_path, url
    return None


def discover_protocols(
    session: requests.Session,
    wp: int,
    cache_dir: Path,
    delay: float,
    max_sessions: int = MAX_SESSIONS_PER_WP,
    max_consecutive_misses: int = MAX_CONSECUTIVE_MISSES,
) -> Iterator[tuple[int, Path, str]]:
    """Probiert Sitzungsnummern sequenziell durch, bis mehrere Treffer in Folge fehlen."""
    misses = 0
    for nr in range(1, max_sessions + 1):
        result = download_protocol(session, wp, nr, cache_dir)
        if result is None:
            misses += 1
            if misses >= max_consecutive_misses:
                break
            continue
        misses = 0
        path, url = result
        yield nr, path, url
        if delay:
            time.sleep(delay)


def extract_pages_text(pdf_path: Path) -> list[str]:
    if pdfplumber is None:
        raise RuntimeError("pdfplumber ist nicht installiert (pip install pdfplumber)")
    pages = []
    with pdfplumber.open(str(pdf_path)) as pdf:
        for page in pdf.pages:
            pages.append(page.extract_text() or "")
    return pages


def extract_datum(first_pages_text: str) -> str:
    match = DATE_RE.search(first_pages_text)
    return match.group(1) if match else ""


def parse_speeches(pages: list[str], wp: int, sitzung: int, quelle_url: str) -> list[Redebeitrag]:
    """Segmentiert den Volltext eines Protokolls in einzelne Redebeitraege."""
    full_text = "\n".join(pages)
    datum = extract_datum("\n".join(pages[:2]))

    lines = full_text.splitlines()
    beitraege: list[Redebeitrag] = []
    current_speaker = None
    current_rolle = ""
    current_buffer: list[str] = []
    seq = 0

    def flush():
        nonlocal seq
        if current_speaker and current_buffer:
            text = " ".join(l.strip() for l in current_buffer if l.strip())
            if text:
                seq += 1
                beitraege.append(
                    Redebeitrag(
                        wahlperiode=wp,
                        sitzung=sitzung,
                        datum=datum,
                        nr=seq,
                        redner=current_speaker,
                        rolle_fraktion=current_rolle,
                        text=text,
                        quelle_url=quelle_url,
                    )
                )

    for raw_line in lines:
        line = raw_line.strip()
        if not line:
            continue
        m = SPEAKER_RE.match(line)
        # Heuristik gegen False Positives: Name-Teil darf nicht zu lang sein, muss
        # aus mindestens zwei Woertern bestehen (Vorname + Nachname, ggf. mit
        # Praesidiums-Titel) und mit Grossbuchstaben beginnen; Zwischenrufe/
        # Regieanweisungen wie "(Beifall bei der CDU)" beginnen mit "(" und werden
        # hier nicht erfasst. Kurze Ablaufvermerke wie "Beginn: 10:00 Uhr" werden
        # durch die Mindestwortanzahl ausgeschlossen.
        if (
            m
            and len(m.group("name")) <= 60
            and len(m.group("name").split()) >= 2
            and not line.startswith("(")
        ):
            flush()
            current_speaker = m.group("name").strip()
            current_rolle = (m.group("rolle") or "").strip()
            current_buffer = [m.group("rest")] if m.group("rest") else []
        elif STAGE_DIRECTION_RE.match(line):
            continue  # Zwischenruf/Regieanweisung, nicht Teil des Redetexts
        else:
            if current_speaker:
                current_buffer.append(line)
    flush()
    return beitraege


def load_checkpoint(path: Path) -> tuple[dict[tuple[int, int], list[Redebeitrag]], list[Redebeitrag]]:
    """Laedt bereits verarbeitete Sitzungen aus einer Checkpoint-Datei (JSON Lines),
    damit ein Neustart nicht wieder von vorn parsen muss."""
    done: dict[tuple[int, int], list[Redebeitrag]] = {}
    rows: list[Redebeitrag] = []
    if not path.exists():
        return done, rows
    with path.open("r", encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if not line:
                continue
            obj = json.loads(line)
            session_rows = [Redebeitrag(**r) for r in obj["rows"]]
            done[(obj["wp"], obj["nr"])] = session_rows
            rows.extend(session_rows)
    return done, rows


def append_checkpoint(path: Path, wp: int, nr: int, rows: list[Redebeitrag]) -> None:
    """Haengt das Ergebnis einer Sitzung an die Checkpoint-Datei an und erzwingt
    sofortiges Schreiben auf die Platte, damit bei einem Absturz nichts verloren geht."""
    entry = {"wp": wp, "nr": nr, "rows": [asdict(r) for r in rows]}
    with path.open("a", encoding="utf-8") as f:
        f.write(json.dumps(entry, ensure_ascii=False))
        f.write("\n")
        f.flush()
        os.fsync(f.fileno())


def write_excel(rows: list[Redebeitrag], out_path: Path) -> None:
    if Workbook is None:
        raise RuntimeError("openpyxl ist nicht installiert (pip install openpyxl)")
    wb = Workbook()
    ws = wb.active
    ws.title = "Redebeitraege"
    headers = ["Wahlperiode", "Sitzung", "Datum", "Nr", "Redner", "Rolle/Fraktion", "Redebeitrag", "Quelle"]
    ws.append(headers)
    for r in rows:
        ws.append([r.wahlperiode, r.sitzung, r.datum, r.nr, r.redner, r.rolle_fraktion, r.text, r.quelle_url])
    ws.freeze_panes = "A2"
    widths = [11, 8, 11, 5, 28, 30, 90, 45]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    out_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(str(out_path))


SELFTEST_SAMPLE = """\
Landtag von Sachsen-Anhalt Stenografischer Bericht 6/42

Beginn: 10:00 Uhr

Präsidentin Gabriele Brakebusch:
Meine sehr geehrten Damen und Herren! Ich eröffne die 42. Sitzung des Landtages
von Sachsen-Anhalt am 22.03.2013 und begrüße Sie alle sehr herzlich.

Guido Kosmehl (FDP):
Sehr geehrte Frau Präsidentin! Liebe Kolleginnen und Kollegen! Wir beraten heute
über einen wichtigen Antrag zur Bildungspolitik in unserem Land.

(Beifall bei der FDP)

Petra Grimm-Benne (Ministerin für Arbeit, Soziales und Integration):
Vielen Dank, Herr Präsident. Die Landesregierung unterstützt diesen Vorschlag
ausdrücklich und wird die entsprechenden Mittel bereitstellen.

Präsidentin Gabriele Brakebusch:
Ich schließe die Aussprache und komme zur Abstimmung.
"""


def run_selftest() -> bool:
    print("[Selbsttest] Parser wird offline gegen ein synthetisches Beispielprotokoll geprueft ...")
    pages = [SELFTEST_SAMPLE]
    beitraege = parse_speeches(pages, wp=6, sitzung=42, quelle_url="(selftest)")

    ok = True
    if len(beitraege) != 4:
        print(f"  [FAIL] Erwartet 4 Redebeitraege, gefunden: {len(beitraege)}")
        ok = False
    else:
        print(f"  [OK] {len(beitraege)} Redebeitraege erkannt")

    expected_speakers = [
        "Präsidentin Gabriele Brakebusch",
        "Guido Kosmehl",
        "Petra Grimm-Benne",
        "Präsidentin Gabriele Brakebusch",
    ]
    for i, (b, expected) in enumerate(zip(beitraege, expected_speakers), start=1):
        if b.redner != expected:
            print(f"  [FAIL] Redebeitrag {i}: Redner '{b.redner}' != erwartet '{expected}'")
            ok = False

    if beitraege and beitraege[1].rolle_fraktion != "FDP":
        print(f"  [FAIL] Fraktion von Redebeitrag 2 sollte 'FDP' sein, ist '{beitraege[1].rolle_fraktion}'")
        ok = False

    if beitraege and "22.03.2013" not in beitraege[0].datum:
        print(f"  [FAIL] Datum nicht korrekt erkannt: '{beitraege[0].datum}'")
        ok = False

    if beitraege and "(Beifall" in beitraege[1].text:
        print("  [FAIL] Zwischenruf '(Beifall ...)' wurde faelschlich in den Redetext uebernommen")
        ok = False

    print("[Selbsttest] " + ("BESTANDEN" if ok else "FEHLGESCHLAGEN"))
    return ok


def run(
    wahlperioden: list[int],
    out_path: Path,
    cache_dir: Path,
    delay: float,
    max_sessions: int,
    test_mode: bool,
    speaker_filter: Optional[str] = None,
    checkpoint_path: Optional[Path] = None,
) -> None:
    session = build_session()
    all_rows: list[Redebeitrag] = []
    done: dict[tuple[int, int], list[Redebeitrag]] = {}
    if checkpoint_path is not None:
        done, checkpoint_rows = load_checkpoint(checkpoint_path)
        all_rows.extend(checkpoint_rows)
        if done:
            print(f"[Checkpoint] {len(checkpoint_rows)} Redebeitraege aus {len(done)} "
                  f"bereits verarbeiteten Sitzungen aus {checkpoint_path} geladen.")

    for wp in wahlperioden:
        print(f"\n=== Wahlperiode {wp}: suche Plenarprotokolle ===")
        found_any = False
        for nr, pdf_path, source in discover_protocols(session, wp, cache_dir, delay, max_sessions=max_sessions):
            found_any = True
            key = (wp, nr)
            if key in done:
                rows = done[key]
                print(f"  Sitzung {nr:03d}: (checkpoint) {len(rows)} Redebeitraege")
            else:
                print(f"  Sitzung {nr:03d}: {source} -> {pdf_path.name}")
                try:
                    pages = extract_pages_text(pdf_path)
                except Exception as exc:
                    print(f"    [WARN] PDF-Extraktion fehlgeschlagen: {exc}", file=sys.stderr)
                    continue
                rows = parse_speeches(pages, wp, nr, source)
                if speaker_filter:
                    rows = [r for r in rows if speaker_filter.lower() in r.redner.lower()]
                print(f"    -> {len(rows)} Redebeitraege extrahiert")
                all_rows.extend(rows)
                if checkpoint_path is not None:
                    append_checkpoint(checkpoint_path, wp, nr, rows)
            if test_mode:
                break  # im Testmodus nur eine Sitzung pro Wahlperiode
        if not found_any:
            print(f"  [WARN] Keine Protokolle fuer Wahlperiode {wp} gefunden "
                  f"(Netzwerk nicht erreichbar oder URL-Schema hat sich geaendert).")

    print(f"\nInsgesamt {len(all_rows)} Redebeitraege gesammelt.")
    if all_rows:
        write_excel(all_rows, out_path)
        print(f"Excel-Datei geschrieben: {out_path}")
    else:
        print("Keine Daten zum Speichern vorhanden - Excel-Datei wurde nicht erzeugt.")


def main() -> None:
    parser = argparse.ArgumentParser(description="Extrahiert Redebeitraege aus Landtag-LSA-Plenarprotokollen.")
    parser.add_argument("--wp", type=int, nargs="+", default=[6, 7, 8], help="Wahlperiode(n), Standard: 6 7 8")
    parser.add_argument("--out", type=Path, default=Path("redebeitraege_lsa.xlsx"), help="Ausgabedatei (.xlsx)")
    parser.add_argument("--cache-dir", type=Path, default=Path("pdfs"), help="Verzeichnis fuer heruntergeladene PDFs")
    parser.add_argument("--delay", type=float, default=0.5, help="Wartezeit zwischen Downloads in Sekunden")
    parser.add_argument("--max-sessions", type=int, default=MAX_SESSIONS_PER_WP, help="Max. Sitzungsnummer pro Wahlperiode")
    parser.add_argument("--test", action="store_true", help="Schneller Funktionscheck: Offline-Selbsttest + 1 Sitzung pro WP")
    parser.add_argument("--speaker", type=str, default=None, help="Nur Redebeitraege dieser Person uebernehmen (Teilstring, Gross-/Kleinschreibung egal)")
    parser.add_argument("--checkpoint", type=Path, default=Path("checkpoint.jsonl"),
                         help="Datei zum laufenden Speichern des Fortschritts (JSON Lines); bei Neustart wird daraus fortgesetzt")
    parser.add_argument("--no-checkpoint", action="store_true", help="Checkpoint-Datei nicht verwenden")
    args = parser.parse_args()

    ok = run_selftest()
    if args.test and not ok:
        print("\nSelbsttest fehlgeschlagen - Abbruch vor dem Netzwerkzugriff.", file=sys.stderr)
        sys.exit(1)

    if args.test:
        print("\n[Testmodus] Lade jeweils nur die erste erreichbare Sitzung pro Wahlperiode ...")
        run(args.wp, Path("test_" + args.out.name), args.cache_dir, args.delay, max_sessions=10, test_mode=True, speaker_filter=args.speaker)
    else:
        checkpoint_path = None if args.no_checkpoint else args.checkpoint
        run(args.wp, args.out, args.cache_dir, args.delay, args.max_sessions, test_mode=False,
            speaker_filter=args.speaker, checkpoint_path=checkpoint_path)


if __name__ == "__main__":
    main()
