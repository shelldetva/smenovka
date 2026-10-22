# Smenovka a odmeny – kompletná dokumentácia a záloha

Stav k: 21. 10. 2026. Projekt pre čerpaciu stanicu Shell Detva.
Tento súbor obsahuje popis systému, jeho prepojenie, rozhodnutia a **kompletné zdrojové kódy** (na konci). Ak stratíš vlákno alebo budeš pokračovať v inej AI, pošli jej tento súbor celý.

---

## 1. Na čo to slúži

Majiteľ plánuje smeny a odmeny v mesačných Excel súboroch. Zamestnanci ich majú vidieť na jednoduchom webe:

- celý rozpis smien (bez prihlásenia, každý s odkazom ho vidí), prepínanie mesiacov,
- odmeny (súťaž medzi zamestnancami, preto sú verejné všetkým),
- aktualizácia sa deje **automaticky** po uložení Excelu,
- rozpis sa dá vytlačiť na **A4 naležato** na nástenku,
- riešenie je **bezplatné** (GitHub Pages).

Web: `https://shelldetva.github.io/smenovka/`
GitHub repozitár: `shelldetva/smenovka` (vetva `main`)

---

## 2. Ako je to poprepájané (architektúra)

```
 Excel (mesačné súbory)                    PC majiteľa (Windows)                    Internet
┌──────────────────────────┐   číta   ┌──────────────────────────────┐  git push  ┌────────────────────┐
│ Master_Smenovka_2026_10  │ ───────► │ sync.py  (beží skryte na      │ ─────────► │ GitHub repo        │
│ .xlsx  (listy MASTER,    │          │ pozadí, spúšťa ho            │            │ shelldetva/        │
│ SUMAR, ...)              │          │ spusti_skryte.vbs)           │            │ smenovka (main)    │
└──────────────────────────┘          │  – prevedie Excel na JSON    │            └─────────┬──────────┘
                                      │  – porovná zmeny             │                      │ GitHub Pages
                                      │  – sám spraví git commit+push│                      ▼
                                      └──────────────────────────────┘            ┌────────────────────┐
                                                                                  │ docs/index.html    │
                                                                                  │ + docs/data/*.json │
                                                                                  │ (web pre ľudí)     │
                                                                                  └────────────────────┘
```

Tok dát: **Excel → sync.py → docs/data/RRRR-MM.json → git push → GitHub Pages → index.html (v prehliadači zamestnanca)**.

Web (`index.html`) je len statická stránka. Nemá server ani databázu. Dáta si načíta z JSON súborov (`data/index.json` a `data/RRRR-MM.json`) a všetko vykreslí JavaScriptom v prehliadači.

---

## 3. Kde je čo (cesty na PC majiteľa)

| Čo | Kde |
|---|---|
| Excely so smenovkou | `C:\Users\Gastro\Desktop\PLANOVANIE A ODMENY NEW` |
| Názov Excelu | presne `Master_Smenovka_RRRR_MM.xlsx` (napr. `Master_Smenovka_2026_10.xlsx`) |
| Projekt (program + web) | `C:\!SHELLAutomatizacia\SmenovkaWEB\` |
| Skript | `C:\!SHELLAutomatizacia\SmenovkaWEB\sync.py` |
| Web | `C:\!SHELLAutomatizacia\SmenovkaWEB\docs\index.html` |
| Dáta pre web (vytvára skript) | `C:\!SHELLAutomatizacia\SmenovkaWEB\docs\data\` |
| Ručné spustenie s oknom | `spusti.bat` |
| Skryté spustenie (používa sa bežne) | `spusti_skryte.vbs` (odkaz naň je v `shell:startup`, spúšťa sa po štarte Windowsu) |
| Výpis skrytého skriptu (log) | `%TEMP%\smenovka_log.txt` |

Súbory v `docs\` sú to, čo GitHub Pages zverejňuje.

---

## 4. Štruktúra Excelu (na čo sa skript spolieha)

Skript **číta Excel iba na čítanie**, vždy z dočasnej kópie, takže nevadí otvorený Excel. Hodnoty číta ako výsledky vzorcov (`data_only=True`), preto musí byť Excel pred čítaním **uložený**.

**List `MASTER`** (nájde sa podľa slova „MASTER“ v názve; skutočný názov v súbore býval `MASTER_SMENOVKA`):
- riadok 1, od stĺpca C: dátumy dní mesiaca,
- od riadku 3: stĺpec A = meno zamestnanca, stĺpec B musí mať text `SMENA` (iné riadky sa preskakujú),
- v bunkách dní sú kódy smien.
- Zamestnanec, ktorý nemá v mesiaci žiadnu smenu (iba „x“ alebo prázdno), sa vynechá.
- Tá istá osoba môže byť v tabuľke pre vnútornú aj vonkajšiu obsluhu (meno vo vonkajšej má „(VO)“).

**List `SUMAR`** (nájde sa podľa slova „SUMAR“):
- blok odmien sa nájde podľa nadpisu v stĺpci B, ktorý začína `VNUTORNA` alebo `VONKAJSIA` (bez diakritiky),
- riadok kategórií = nadpis + 3, dáta zamestnancov od nadpis + 5, až kým je v stĺpci B meno,
- stĺpec B = meno, C = poradie, D = plnenie (podiel, napr. 0,92), E = bonus celkom v €,
- kategórie (napr. cigarety) začínajú od stĺpca H (8) a každá zaberá 5 stĺpcov,
- zamestnanci s nulovým bonusom aj plnením sa nezobrazia (brigádnici).

**List `PLANY`** (nájde sa podľa slova „PLANY“) – **denný plán tovaru podľa typu smeny**:
- bunky so začiatkom `---` sú nadpisy sekcií (napr. `--- 2. VNÚTORNÁ DENNÁ SMENA (D) ---`) a určujú typ smeny: `D-VO` (vonkajšia), `D` (vnútorná denná), `N` (vnútorná nočná),
- stĺpec A = tovar, stĺpec B = merná jednotka (`ks`, `litrov`), stĺpec C = denný plán,
- nočné tovary majú v názve príponu `(noc)` – pre porovnanie s kategóriami v `SUMAR` sa prípona zahodí (`Cigarety (noc)` → `Cigarety`).

| Typ smeny | Denný plán (ríkober 2026) |
|---|---|
| `D-VO` vonkajšia | V-POWER 1550 litrov, SCREENWASH 6 ks, OLEJ 1 ks, AD BLUE 150 litrov |
| `D` vnútorná denná | Cigarety 150 ks, Tabak 32, Doplnky 30, Hotdog 25, Bagety Detvan 22, Káva 23 |
| `N` vnútorná nočná | Cigarety 90 ks, Tabak 36, Doplnky 30, Hotdog 20, Bagety Detvan 22, Káva 6 |

### Kódy smien

| Kód | Význam | Farba na webe |
|---|---|---|
| `D` | denná 06:00–18:00 | žltá |
| `D-VO` | denná vonku 06:00–18:00 | zelená |
| `N` | nočná 18:00–06:00 | modrofialová |
| `POL` | polnočná 18:00–23:00 | svetlomodrá |
| `DOV` | dovolenka | oranžová |
| `PN` | práce neschopnosti | ružová |
| `X` | zamestnanec nemôže (nepozýva sa do práce) | šedá, zamestnanci ho smú vidieť |

Kódy a časy sú oficiálne v Exceli na liste `NASTAVENIA` (A = kód, B = názov, C = čas, D = hodiny). Kde je kód v zátvorkách, ide o doplnkový kód (`PN` = práce neschopnosť).

V bunkách sa môže objaviť aj kód s časom (napr. `POL 18:00-22:30`, `D 06:00-23:00`). Web zobrazí iba samotný kód (`POL`, `D`) a celý kód nechá v názve bunky (po nabehnutí myšou), aby sa tabuľka na papieri nerozpadla.

Dôležité poradie v kóde (`pill()` v `index.html`): kontrola dlhších kódov (`D-VO`, `DOV`, `POL`) musí byť **pred** `D`, inak by sa dovolenka zobrazila ako denná smena.

---

## 5. Formát dát (JSON)

`docs/data/index.json`: zoznam mesiacov na prepínanie.

```json
{"months":[{"id":"2026-10","label":"Október 2026"}]}
```

`docs/data/2026-10.json` (jeden súbor na mesiac):

```
id, label, y (rok), m (mesiac 0–11), days [1..31],
shifts: [ { n: "Meno", s: { "1": "D", "2": "N", ... } } ],
vno / vo: [ { n, rank, tot (bonus €), pl (plnenie podiel),
              c: [ [kategória, plán, skutočnosť, podiel, €, ďalšie], ... ] } ],
plans: { "D-VO": { "V-POWER": { p: 1550, u: "litrov" } },
         "D":    { "CIGARETY": { p: 150, u: "ks" } },
         "N":    { "CIGARETY": { p: 90,  u: "ks" } } },
updated: čas poslednej zmeny,
changes: { "Meno": { "17": { was: "N", at: "ISO čas" } } }
```

Web zobrazuje v kategórii: `c[1]` = plán, `c[2]` = skutočnosť, `c[3]` = podiel, `c[4]` = bonus v €.

`plans` je **denný plán tovaru podľa typu smeny** z listu `PLANY` v Exceli (kľúč = názov kategórie bez diakritiky, veľkými písmenami; `p` = denný plán, `u` = jednotka). Bez neho web vie iba „koľko chýba celkovo“, nie „koľko na konkrétny typ smeny“.

---

## 6. Čo robí `sync.py`

- `EXCEL_FOLDER`, `PUSH`, `PUSH_EVERY_MINUTES` (0 = ihneď), `POLL_SECONDS` (20 s) sú nastavenia na vrchu súboru.
- Každých 20 s skontroluje čas zmeny Excelov. Počká 5 s po poslednom uložení (kým Excel dopíše súbor).
- Spracuje len mesiace: **minulý, aktuálny a nasledujúce 2** (`valid_range`).
- `read_schedule` / `read_rewards` / `read_plans` / `read_month` prevedú Excel na JSON.
- `read_plans()` číta list `PLANY` a uloží **denný plán tovaru podľa typu smeny** (`D`, `N`, `D-VO`) do JSON pod kľúčom `plans` (kategórie bez diakritiky, prípona `(noc)` zahodená).
- Ak sa obsah zmenil (porovnáva sa odtlačok bez `updated` a `changes`), doplní `updated` a `changes` a súbor prepíše.
- `merge_changes` zaznamenáva zmenené smeny **dnes a v budúcnosti** s pôvodnou hodnotou; záznam žije 48 hodín. Nový mesiac a nový zamestnanec sa neoznačujú.
- `push()` urobí `git add -A`, `git commit`, `git push`. Ak nie je čo commitnúť, nevadí.
- `python sync.py --once` spracuje všetko raz a skončí (na skúšku).

---

## 7. Čo robí `index.html`

- Načíta `data/index.json`, zobrazí tlačidlá mesiacov (predvolený je aktuálny).
- **Panel „Dnes a zajtra“** (karta hore): načíta si vlastný JSON aktuálneho mesiaca, takže ukazuje dnešné a zajtrajšie smeny **nezávisle od toho, ktorý mesiac máš práve zvolený**. Stĺpce sú „Dnes“ a „Zajtra“ s dátumom a dňom v týždni; pri každom zamestnancovi je farebná pilulka smeny (`–` = nemá smenu).
- Tabuľka smien: stĺpec s menom pripnutý vľavo, záhlavie dňa má číslo a deň v týždni, víkendy tónované, **minulé dni sivé**, dnešok má žlté záhlavie.
- **Moje smeny:** klik na meno zamestnanca zobrazí len jeho riadok (pamätá si to `localStorage` v jeho prehliadači), pruh hore s tlačidlom „Zobraziť všetkých“.
- **Klik na deň:** kliknutím na číslo dňa sa pod tabuľkou zobrazí „kto ten deň pracuje“ – chipy s menom a smenou; tí, ktorí majú dovolenku / práce neschopnosti / sú nedostupní, sú oddelení nižšie. Klik na ten istý deň panel znova zavrie.
- **Označenie zmien:** červená bodka v bunke + okno „Nedávne zmeny v rozpise“ (48 h, len dnes a neskôr).
- **Odmeny:** tabuľky pre vnútornú a vonkajšiu obsluhu (poradie, plnenie, bonus). Po kliknutí na meno sa rozbalí rozpis kategórií. Ak je vybrané „moje meno“, vlastný riadok v odmenách je zvýraznený žltou farbou.
- **Koľko ešte predať:** pri každej kategórii (napr. cigarety) ukáže, koľko treba predať **na každú zostávajúcu smenu – zvlášť pre každý typ smeny** (denná / nočná / denná vonku), aby zamestnanec dosiahol 90 % a 100 % mesačného plánu. Výpočet: *plán zostávajúcich smien = denný plán z Excelu (`PLANY`) × počet smien daného typu*; deficit (mesačný plán − skutočnosť) sa rozdelí medzi typy úmerne ich plánu a pripočíta sa k plánu týchto smien (zaokrúhlene nahor). Smeny sa počítajú od dneška, bez `X`, `DOV`, `PN` a bez krátkych `POL` (tie nemajú vlastný plán); pri vnútornej obsluhe sa nepočítajú denné smeny vonku `D-VO`, pri vonkajšej naopak. Ak je plán splnený, ukáže zelené potvrdenie a denný plán na zostávajúce smeny. Jednotka (`ks`, `litrov`) sa berie z Excelu a zobrazuje sa.
- **Tlač A4 naležato** (`@media print`): vytlačí len nadpis, legendu a rozpis, bez posuvníkov, odmien, panelu „Dnes a zajtra“, dňového detailu a ovládacích prvkov. Sivé dni ostávajú sivé. Na vytlačenie nástenky nemá byť vybraté žiadne meno (vytlačil by sa len jeden riadok).
- Meno z rozpisu sa páruje s menom v odmenách podľa mena bez diakritiky a bez zátvorky `(VO)`.

---

## 8. Spustenie a nasadenie (postup)

1. Nainštalovaný Python, knižnica `openpyxl` (`pip install openpyxl`), Git s prihlásením na GitHub (`shelldetva`).
2. Skript v priečinku projektu spusti ručne `python sync.py --once` na skúšku.
3. Pre bežnú prevádzku: odkaz na `spusti_skryte.vbs` vlož do `shell:startup` (Win + R). Beží skryto na pozadí, po štarte PC sa spustí samo.
4. Kontrola, či beží: Správca úloh → Podrobnosti → `python.exe`. Log: `%TEMP%\smenovka_log.txt`. Zastavenie: ukončiť `python.exe`.
5. Po zmene kódu skript reštartovať (ukončiť `python.exe` a znova spustiť `.vbs`).
6. Po úprave `index.html`: `git add -A`, `git commit -m "popis"`, `git push`, po 1–2 minútach na webe `Ctrl + F5`.

---

## 9. Problémy, ktoré sme už riešili

| Príznak | Príčina a riešenie |
|---|---|
| Prázdna stránka na webe | Súbor `index.html` sa pri kopírovaní uložil neúplný. Skopíruj súbor, nie text. |
| „Chyba načítania zoznamu mesiacov“ pri otvorení `index.html` z disku | Prehliadač blokuje `data/*.json` zo súboru na disku. Otváraj cez `https://shelldetva.github.io/smenovka/`. Na webe to funguje. |
| Zmena z Excelu sa neukáže | Skript nebeží, Excel nebol uložený, GitHub Pages potrebuje 1–2 min, alebo je mesiac mimo rozsah (-1 až +2). |
| `spusti.bat` skončí hneď s „Press any key“ a nič nevypíše | `sync.py` bol useknutý (chýbal koniec s `main()`). Prepísať celým súborom. |
| Zvláštne znaky v okne príkazového riadka | Len zobrazenie v konzole, v súbore je text v poriadku. |
| Dovolenka sa zobrazovala ako denná | Poradie kontrol kódov v `pill()` (opravené). |
| `PN` (práce neschopnosti) sa zobrazoval bez farby a bez legendy | Web poznal len D, D-VO, N, POL, DOV, X. Doplnený kód `PN` (ružová) aj do legendy a do počítania zostávajúcich smien. |
| Kódy s časom (`POL 18:00-22:30`) roztláčali tabuľku | Web zobrazuje iba samotný kód, celý kód ostáva v názve bunky. |
| „Koľko ešte predať“ delilo rovnaký rozdiel na všetky smeny | Nočná smena má iný denný plán ako denná (cigarety 90 vs. 150) a vonkajšia úplne iný (V-POWER 1550 litrov). Výpočet prepočítaný podľa typu smeny z listu `PLANY`. |
| Bodky zmien sa neukazujú | Bežala stará verzia `sync.py` bez `merge_changes`. Prepísať a reštartovať skript. |

Dva skripty naraz sa nemajú spúšťať (`spusti.bat` a `spusti_skryte.vbs` súčasne).

---

## 10. Rozhodnutia a čo sa zámerne nerobí

- Zamestnanci vidia **celý rozpis aj červené „X“** aj všetky odmeny (súťaž). Žiadne prihlasovanie.
- Zamestnanci si smeny medzi sebou občas vymieňajú a oznámia to majiteľovi. Rozpis sa upravuje len v Exceli.
- Export do kalendára (`.ics`) sa **nebude robiť**.
- Hosting: GitHub Pages (pôvodne zvažovaný Cloudflare Pages).
- Majiteľ si to chce vedieť spravovať sám.

---

## 11. Známe obmedzenia a čo overiť

- Mesačný plán v `SUMAR` (a teda aj na webe) je súčet denných plánov len za **dni, ktoré už majú zapísané predaje** – tak to počíta majiteľov Excel. Preto sa „koľko ešte predať“ počíta ako *deficit (mesačný plán − skutočnosť) + denný plán zostávajúcich smien*, a nie jednoduchým delením rozdielu.
- Výpočet „koľko ešte predať“ potrebuje v JSON kľúč `plans` (list `PLANY` v Exceli). Ak v mesiaci chýba, web ukáže iba celkový rozdiel „do 90 % / do 100 % chýba X“.
- Dnešný deň sa pri výpočte smien počíta ako ešte neodpracovaný.
- Do počtu zostávajúcich smien sa **nezapočítavajú** krátke polnočné smeny `POL` (nemajú v `PLANY` vlastný plán) a (pri vnútornej obsluhe) denné smeny vonku `D-VO` – je to pravidlo majiteľa, mení sa v `shiftsByType()`.
- Ak sa meno v `SUMAR` píše inak ako v `MASTER`, počet smien sa nezobrazí (web vtedy ukáže len celkový rozdiel).
- Pri zamestnancovi, ktorý robí vnútornú aj vonkajšiu obsluhu (`VO/VNO`), počíta Excel cigarety aj z dní `D-VO` (v jeho liste je iný vzorec než u ostatných). Web pre vnútornú obsluhu počíta len `D` a `N`, pre vonkajšiu len `D-VO` – pri zostávajúcich `D-VO` smenách sa preto môže plán cigariet mierne líšiť od Excelu.
- Ak niekto v Exceli opraví preklep, označí sa to ako zmena smeny.
- Nástenku tlačiť bez vybraného mena.
- Výber „moje meno“ je len v prehliadači jednotlivého zamestnanca (nie je to zabezpečenie, dáta sú verejné).

---

## 12. Nápady na ďalší vývoj

**Už hotovo (21. 10. 2026):** panel „Dnes a zajtra“, klik na deň (kto pracuje), kód `PN` v legende, skrátenie kódov s časom, zvýraznenie vlastného riadka v odmenách.

**Už hotovo (22. 10. 2026):** „koľko ešte predať“ sa počíta podľa typu zostávajúcej smeny (denná / nočná / denná vonku) z nového listu `PLANY`; jednotky (`ks`, `litrov`) sa zobrazujú.

- QR kód na vytlačenom rozpise s odkazom na web.
- Oznam od vedúceho z bunky v Exceli (žltý pruh na webe).
- Rebríček s medailami, pásik plnenia, rozdiel k prvému miestu, šípky pohybu v poradí, víťaz mesiaca.
- Tmavý režim, ikona na ploche mobilu (PWA).
- Prípadne skryť sumy v € a nechať v súťaži len poradie a percentá.

---

## 13. Pokyn pre inú AI (vlož na začiatok rozhovoru)

> Pokračujeme vo vývoji webu so smenovkou a odmenami pre čerpaciu stanicu. Pozri si priloženú dokumentáciu a zdrojové kódy na konci. Systém: Excel → `sync.py` (Python, openpyxl, beží na Windows PC) → JSON do `docs/data` → git push → GitHub Pages → `docs/index.html` (statická stránka, JavaScript bez knižníc). Používateľ nie je programátor, píše po slovensky, potrebuje jednoduché kroky po jednom. Pri zmenách vždy dodaj **kompletné** súbory (nie len úryvky), aby sa nepoškodili pri kopírovaní. Pred úpravou over, či sedí štruktúra Excelu (kapitola 4). Zachovaj A4 tlač, sivé minulé dni a verejné odmeny.
>
> Nové funkcie (21. 10. 2026) sú v `docs/index.html`; 22. 10. 2026 pribudol do `sync.py` list `PLANY` (`plans` v JSON) pre výpočet „koľko ešte predať“ podľa typu smeny. Ak pridávaš novú funkciu, najprv ju odtestuj lokálne: v priečinku `docs` spusti `python -m http.server 8765` a otvor `http://127.0.0.1:8765/` (z disku cez `file://` prehliadač blokuje načítanie `data/*.json`).

---

## 14. Zdrojové kódy (aktuálne verzie)

### spusti.bat (ručné spustenie s oknom)

```bat
@echo off
cd /d "%~dp0"
python sync.py
pause
```

### spusti_skryte.vbs (skryté spustenie, dať odkaz do shell:startup)

```vbs
' Spustí sync.py úplne skryte (bez okna). Výpis ukladá do %TEMP%\smenovka_log.txt
Set s = CreateObject("WScript.Shell")
s.CurrentDirectory = "C:\!SHELLAutomatizacia\SmenovkaWEB"
s.Run "cmd /c python -u sync.py > ""%TEMP%\smenovka_log.txt"" 2>&1", 0, False
```

### sync.py

```python
"""Prevod mesačných Excelov so smenovkou na dáta pre web.

Spustenie:
    python sync.py          sleduje priečinok s Excelmi a po každom uložení aktualizuje web
    python sync.py --once   spracuje všetko raz a skončí
"""
import datetime
import hashlib
import json
import re
import shutil
import subprocess
import sys
import tempfile
import time
import unicodedata
from pathlib import Path

from openpyxl import load_workbook

# ------------------------- NASTAVENIA -------------------------
EXCEL_FOLDER = Path(r"C:\Users\Gastro\Desktop\PLANOVANIE A ODMENY NEW")
PUSH = True                               # True = po zmene sám nahrá web na GitHub
PUSH_EVERY_MINUTES = 0                    # 0 = odosiela zmeny ihneď po uložení
POLL_SECONDS = 20                         # ako často kontroluje priečinok
# --------------------------------------------------------------

BASE = Path(__file__).resolve().parent
SITE = BASE / "docs"
DATA = SITE / "data"
FILE_RE = re.compile(r"^Master_Smenovka_(\d{4})_(\d{2})\.xlsx$", re.I)
MONTHS = ["Január", "Február", "Marec", "Apríl", "Máj", "Jún", "Júl",
          "August", "September", "Október", "November", "December"]


def plain(v):
    """Text bez diakritiky, veľkými písmenami (na porovnávanie)."""
    s = unicodedata.normalize("NFKD", str(v or ""))
    return "".join(c for c in s if not unicodedata.combining(c)).strip().upper()


def find_sheet(wb, *words):
    for ws in wb.worksheets:
        if all(w in plain(ws.title) for w in words):
            return ws
    raise ValueError("V súbore chýba list: " + " ".join(words))


def read_schedule(ws, year, month):
    cols = {}
    for c in range(3, ws.max_column + 1):
        v = ws.cell(1, c).value
        if isinstance(v, datetime.datetime) and v.year == year and v.month == month:
            cols[v.day] = c
    days = sorted(cols)
    shifts = []
    for r in range(3, ws.max_row + 1):
        name = ws.cell(r, 1).value
        if not name or plain(ws.cell(r, 2).value) != "SMENA":
            continue
        s = {}
        for d in days:
            v = ws.cell(r, cols[d]).value
            if v not in (None, ""):
                s[str(d)] = str(v).strip()
        # pracovníci, ktorí nemajú v mesiaci žiadnu smenu (len "x" alebo prázdno), sa vynechajú
        if any(not code.upper().startswith("X") for code in s.values()):
            shifts.append({"n": str(name).strip(), "s": s})
    return days, shifts


def read_rewards(ws, key):
    """Nájde blok odmien podľa nadpisu v stĺpci B (napr. VNUTORNA OBSLUHA)."""
    marker = next((r for r in range(1, ws.max_row + 1)
                   if plain(ws.cell(r, 2).value).startswith(key)), None)
    if marker is None:
        return []
    cat_row = marker + 3
    cats = [(str(ws.cell(cat_row, c).value).strip(), c)
            for c in range(8, ws.max_column + 1) if ws.cell(cat_row, c).value]
    out = []
    r = marker + 5
    while r <= ws.max_row and ws.cell(r, 2).value:
        name = str(ws.cell(r, 2).value).strip()
        num = lambda row, col: round(float(ws.cell(row, col).value or 0), 4)
        total, fulfil = num(r, 5), num(r, 4)
        if total or fulfil:                       # brigádnici s nulami sa nezobrazia
            rank = ws.cell(r, 3).value
            out.append({
                "n": name,
                "rank": rank if isinstance(rank, (int, float)) else None,
                "tot": round(total, 2), "pl": fulfil,
                "c": [[cn.title() if cn.isupper() else cn,
                       num(r, c), num(r, c + 1), num(r, c + 2),
                       round(num(r, c + 3), 2), round(num(r, c + 4), 2)]
                      for cn, c in cats],
            })
        r += 1
    out.sort(key=lambda x: (x["rank"] is None, x["rank"] or 0))
    return out


def base_cat(name):
    """Názov kategórie bez zátvorky na konci ('Cigarety (noc)' -> 'Cigarety')."""
    return re.sub(r"\s*\(.*?\)\s*$", "", str(name)).strip()


def read_plans(ws):
    """Denný plán tovaru podľa typu smeny (D-VO / D / N) z listu PLANY.

    List PLANY má sekcie:
        --- 1. VONKAJŠIA SMENA (D-VO) ---   stĺpec A = tovar, B = merná jednotka, C = denný plán
        --- 2. VNÚTORNÁ DENNÁ SMENA (D) ---
        --- 3. VNÚTORNÁ NOČNÁ SMENA (N) ---
    Vrací {"D": {"CIGARETY": {"p": 150, "u": "ks"}}, "N": {...}, "D-VO": {...}}
    """
    plans = {}
    kind = None
    for r in range(1, ws.max_row + 1):
        a = ws.cell(r, 1).value
        if a in (None, ""):
            continue
        s = str(a).strip()
        if s.startswith("---"):                      # nadpis sekcie
            up = plain(s)
            if "D-VO" in up:
                kind = "D-VO"
            elif "NOC" in up:
                kind = "N"
            elif "DENN" in up:
                kind = "D"
            else:
                kind = None
            continue
        if not kind:
            continue
        val = ws.cell(r, 3).value
        if not isinstance(val, (int, float)):
            continue
        cat = base_cat(s)
        if not cat:
            continue
        plans.setdefault(kind, {})[plain(cat)] = {
            "p": round(float(val), 4),
            "u": str(ws.cell(r, 2).value or "").strip(),
        }
    return plans


def read_month(path, year, month):
    with tempfile.TemporaryDirectory() as tmp:    # kópia, aby nevadil otvorený Excel
        copy = Path(tmp) / "x.xlsx"
        shutil.copy2(path, copy)
        wb = load_workbook(copy, data_only=True)
    days, shifts = read_schedule(find_sheet(wb, "MASTER"), year, month)
    su = find_sheet(wb, "SUMAR")
    try:
        plans = read_plans(find_sheet(wb, "PLANY"))
    except Exception:
        plans = {}
    return {
        "id": f"{year}-{month:02d}", "label": f"{MONTHS[month - 1]} {year}",
        "y": year, "m": month - 1, "days": days, "shifts": shifts,
        "vno": read_rewards(su, "VNUTORNA"), "vo": read_rewards(su, "VONKAJSIA"),
        "plans": plans,
    }


def digest(d):
    return hashlib.md5(json.dumps(d, sort_keys=True, ensure_ascii=False).encode()).hexdigest()


CHANGE_HOURS = 48                         # ako dlho sa zmena na webe zvýrazňuje


def merge_changes(old_full, data, today):
    """Zmenené smeny (dnes a v budúcnosti) + pôvodná hodnota; staré záznamy po 48 h zmiznú."""
    now = datetime.datetime.now().astimezone()
    limit = datetime.timedelta(hours=CHANGE_HOURS)
    ch = {}
    for name, days in (old_full.get("changes") or {}).items():
        for day, info in days.items():
            try:
                fresh = now - datetime.datetime.fromisoformat(info["at"]) < limit
                future = datetime.date(data["y"], data["m"] + 1, int(day)) >= today
            except Exception:
                continue
            if fresh and future:
                ch.setdefault(name, {})[day] = info
    old_map = {e["n"]: e["s"] for e in old_full.get("shifts", [])}
    if old_map:                                   # nový mesiac sa neoznačuje
        for e in data["shifts"]:
            if e["n"] not in old_map:             # nový zamestnanec sa neoznačuje
                continue
            for d in data["days"]:
                new, prev = e["s"].get(str(d), ""), old_map[e["n"]].get(str(d), "")
                if new != prev and datetime.date(data["y"], data["m"] + 1, d) >= today:
                    ch.setdefault(e["n"], {})[str(d)] = {"was": prev or "-", "at": now.isoformat(timespec="seconds")}
    return ch


def process_all():
    DATA.mkdir(parents=True, exist_ok=True)
    changed, months = False, []

    # Aktuálny rok a mesiac: 1 mesiac dozadu až 2 mesiace dopredu (-1, 0, +1, +2)
    today = datetime.date.today()
    curr_idx = today.year * 12 + (today.month - 1)
    valid_range = {curr_idx - 1, curr_idx, curr_idx + 1, curr_idx + 2}

    for f in sorted(EXCEL_FOLDER.glob("Master_Smenovka_*.xlsx")):
        m = FILE_RE.match(f.name)
        if not m:
            continue
        year, month = int(m.group(1)), int(m.group(2))
        f_idx = year * 12 + (month - 1)

        # Vyfiltruj iba povolené mesiace (-1 až +2)
        if f_idx not in valid_range:
            continue

        data = read_month(f, year, month)
        months.append({"id": data["id"], "label": data["label"]})
        target = DATA / f"{data['id']}.json"
        old_full = json.loads(target.read_text("utf-8")) if target.exists() else {}
        old = dict(old_full)
        old.pop("updated", None)
        old.pop("changes", None)
        if digest(old) != digest(data):
            data["changes"] = merge_changes(old_full, data, today)
            data["updated"] = datetime.datetime.now().astimezone().isoformat(timespec="seconds")
            target.write_text(json.dumps(data, ensure_ascii=False, separators=(",", ":")), "utf-8")
            print(time.strftime("%H:%M:%S"), "aktualizované:", f.name)
            changed = True

    index = DATA / "index.json"
    new_index = json.dumps({"months": months}, ensure_ascii=False)
    if not index.exists() or index.read_text("utf-8") != new_index:
        index.write_text(new_index, "utf-8")
        changed = True
    return changed


def push():
    for cmd in (["git", "add", "-A"],
                ["git", "commit", "-m", "aktualizácia smenovky"],
                ["git", "push"]):
        r = subprocess.run(cmd, cwd=BASE, capture_output=True, text=True)
        if r.returncode and "nothing to commit" not in (r.stdout + r.stderr):
            print("Git zlyhal:", (r.stderr or r.stdout).strip())
            return False
    print(time.strftime("%H:%M:%S"), "web nahraný, na adrese bude za minútu-dve")
    return True


def signature():
    return tuple((p.name, p.stat().st_mtime) for p in sorted(EXCEL_FOLDER.glob("Master_Smenovka_*.xlsx")))


def main():
    once = "--once" in sys.argv
    last, pending, last_push = None, False, 0.0
    print("Sledujem priečinok:", EXCEL_FOLDER)
    while True:
        try:
            sig = signature()
            settled = all(time.time() - mt > 5 for _, mt in sig)   # počká na dokončenie ukladania v Exceli
            if sig != last and settled:
                if process_all():
                    pending = True
                last = sig
            due = time.time() - last_push >= PUSH_EVERY_MINUTES * 60
            if pending and PUSH and (due or once):
                if push():
                    pending, last_push = False, time.time()
        except Exception as e:
            print(time.strftime("%H:%M:%S"), "čakám:", e)
        if once:
            break
        time.sleep(POLL_SECONDS)


if __name__ == "__main__":
    main()
```

### docs/index.html

```html
<!DOCTYPE html>
<html lang="sk">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Smenovka a odmeny</title>
<style>
:root{
  --ink:#16181d; --muted:#6b7280; --line:#e5e7eb; --paper:#f6f5f2; --card:#fff;
  --brand:#ffc600; --brand-ink:#3b2f00; --red:#dd1d21;
  --past-bg:#eceef0; --past-ink:#a4abb3;
}
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:"Segoe UI Variable","Segoe UI",system-ui,-apple-system,Roboto,Arial,sans-serif;
  background:var(--paper);color:var(--ink);padding:24px;-webkit-print-color-adjust:exact;print-color-adjust:exact}
.wrap{max-width:1500px;margin:0 auto}
header{display:flex;justify-content:space-between;align-items:flex-end;flex-wrap:wrap;gap:14px;margin-bottom:16px;
  border-left:6px solid var(--brand);padding-left:14px}
h1{font-size:28px;font-weight:800;letter-spacing:-.02em;line-height:1.1}
.sub{font-size:13px;color:var(--muted);margin-top:4px}
.controls{display:flex;gap:8px;flex-wrap:wrap;align-items:center}
#month-list{display:flex;gap:6px;flex-wrap:wrap}
button{font:inherit;cursor:pointer}
.month-btn,.print-btn{border:1px solid #d4d7dc;background:var(--card);color:#374151;padding:8px 16px;border-radius:10px;font-size:13px;font-weight:600}
.month-btn:hover,.print-btn:hover{background:#f0f1f3}
.month-btn.active{background:var(--ink);border-color:var(--ink);color:#fff}
.print-btn{background:var(--brand);border-color:var(--brand);color:var(--brand-ink)}
.print-btn:hover{background:#f0b900}
button:focus-visible{outline:3px solid var(--red);outline-offset:2px}

.legend{display:flex;gap:14px;flex-wrap:wrap;font-size:12px;color:#4b5563;margin-bottom:14px}
.legend span{display:inline-flex;align-items:center;gap:6px}
.pill{display:inline-block;min-width:26px;padding:3px 6px;border-radius:6px;font-weight:700;font-size:11px;line-height:1.2}
.p-d{background:#ffe27a;color:#6b4e00}
.p-dvo{background:#bff0cf;color:#14532d}
.p-n{background:#c7d2fe;color:#312e81}
.p-pol{background:#b6e3f7;color:#0c4a6e}
.p-dov{background:#fdc9a0;color:#7c2d12}
.p-pn{background:#fecdd3;color:#881337}
.p-x{background:#f1f2f4;color:#9aa1ab}

.card{background:var(--card);border:1px solid var(--line);border-radius:16px;margin-bottom:22px;overflow:hidden}
.card-h{padding:14px 18px;font-size:16px;font-weight:700;border-bottom:1px solid var(--line);display:flex;justify-content:space-between;gap:10px;flex-wrap:wrap}
.card-h small{font-weight:400;font-size:12px;color:var(--muted)}
.tw{overflow-x:auto}
table{width:100%;border-collapse:separate;border-spacing:0;text-align:center;font-size:12px}
th,td{border-right:1px solid var(--line);border-bottom:1px solid var(--line);padding:5px 2px;min-width:30px}
th{background:#fafafa;font-weight:600;color:#4b5563;line-height:1.25}
th .dn{display:block;font-size:13px;font-weight:700;color:var(--ink)}
th .dw{display:block;font-size:10px;font-weight:500;color:var(--muted)}
.emp{text-align:left;font-weight:600;min-width:170px;max-width:170px;padding-left:14px;position:sticky;left:0;background:#fff;z-index:2;border-right:2px solid #d4d7dc;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
th.emp{background:#fafafa}
.we{background:#faf9f6}
th.we{background:#f3f1ea}
th.today{background:var(--brand)}
th.today .dn,th.today .dw{color:var(--brand-ink)}
th.day{cursor:pointer}
th.day:hover{background:#eef0f3}
th.sel{outline:3px solid var(--brand);outline-offset:-3px}
#body .emp{cursor:pointer}
#body tr:hover .emp{text-decoration:underline}
.mebar{display:none;align-items:center;justify-content:space-between;gap:10px;flex-wrap:wrap;padding:10px 18px;background:#fff8d6;border-bottom:1px solid var(--line);font-size:13px}
.mebar button{border:1px solid #d4d7dc;background:#fff;border-radius:8px;padding:5px 12px;font-size:12px;font-weight:600}
td.chg{position:relative}
td.chg::after{content:"";position:absolute;top:2px;right:2px;width:7px;height:7px;border-radius:50%;background:var(--red)}
.chgbox{display:none}
.chgbox ul{list-style:none;padding:12px 18px}
.chgbox li{padding:4px 0;font-size:13px}
.chgbox .dot{display:inline-block;width:8px;height:8px;border-radius:50%;background:var(--red);margin-right:8px}
/* uplynulé dni = sivé */
td.past,th.past{background:var(--past-bg)!important;color:var(--past-ink)!important}
th.past .dn,th.past .dw{color:var(--past-ink)!important}
.past .pill{background:#dcdfe3!important;color:var(--past-ink)!important}
tbody tr:last-child td{border-bottom:none}
.foot{padding:10px 18px;font-size:12px;color:var(--muted);border-top:1px solid var(--line)}

/* --- panel Dnes a zajtra --- */
.todaybox td{padding:7px 6px}
.todaybox th.emp{width:170px}
.todaybox .none{color:var(--muted)}
.todaybox .wd{font-size:10px;font-weight:400;color:var(--muted)}
.todaybox td.emp{position:static;min-width:170px;max-width:170px}

/* --- klik na deň: kto pracuje --- */
.daybox{display:none;border-top:1px solid var(--line);background:#fbfbf9;padding:12px 18px;font-size:13px}
.daybox .dh{font-weight:700;margin-bottom:8px;font-size:14px}
.daybox .dl{display:flex;flex-wrap:wrap;gap:6px 10px}
.daybox .one{display:inline-flex;align-items:center;gap:6px;background:#fff;border:1px solid var(--line);border-radius:8px;padding:4px 9px}
.daybox .one b{font-weight:600}
.daybox .one.off{background:#f7f7f5;color:#7a8089}
.daybox .none{color:var(--muted)}

.rw{cursor:pointer}.rw:hover td{background:#fafafa}
.rw td.emp{background:inherit}
.rw.mine td{background:#fff8d6}
.rw.mine td.emp{background:#fff8d6;font-weight:700}
.dt{display:none;background:#fafafa}
.dt td{text-align:left;padding:10px 14px}
.dt.mine{background:#fff8d6}
.gap{display:block;margin-top:4px;font-size:12px;color:#92400e}
.ok{color:#15803d;font-weight:700}
.chip{display:inline-block;margin:3px 6px 3px 0;padding:5px 9px;border-radius:8px;background:#fff;border:1px solid var(--line);font-size:12px}

@media (max-width:700px){body{padding:12px}h1{font-size:22px}}

@media print{
  @page{size:A4 landscape;margin:6mm}
  html,body{background:#fff!important;padding:0!important;color:#000}
  .controls,.rewards,.foot,.chgbox,.mebar,.todaybox,.daybox{display:none!important}
  .wrap{max-width:none}
  header{margin-bottom:6px;padding-left:8px;border-left-width:4px}
  h1{font-size:16pt}
  .sub{font-size:8pt;margin-top:1px}
  .legend{font-size:7pt;gap:10px;margin-bottom:5px}
  .pill{font-size:6.5pt;padding:1px 3px;min-width:0}
  .card{border:none;border-radius:0;margin:0;overflow:visible}
  .card-h{padding:0 0 4px;border:none;font-size:10pt}
  .tw{overflow:visible!important}
  table{table-layout:fixed;font-size:7pt;border-collapse:collapse}
  th,td{border:1px solid #cfd3d8;padding:2.2mm 0;min-width:0}
  th.sel{outline:none!important}
  th{padding:1mm 0}
  th .dn{font-size:7.5pt}th .dw{font-size:6pt}
  .emp{position:static;width:34mm;min-width:34mm;max-width:34mm;font-size:7.5pt;padding-left:1.5mm;border-right:1.5px solid #9ca3af}
}
</style>
</head>
<body>
<div class="wrap">
  <header>
    <div>
      <h1>Smenovka a odmeny</h1>
      <div class="sub" id="upd"></div>
    </div>
    <div class="controls">
      <div id="month-list"></div>
      <button class="print-btn" onclick="window.print()">🖨️ Tlačiť na A4</button>
    </div>
  </header>

  <div class="legend">
    <span><i class="pill p-d">D</i> denná 06:00–18:00</span>
    <span><i class="pill p-dvo">D-VO</i> denná vonku 06:00–18:00</span>
    <span><i class="pill p-n">N</i> nočná 18:00–06:00</span>
    <span><i class="pill p-pol">POL</i> polnočná 18:00–23:00</span>
    <span><i class="pill p-dov">DOV</i> dovolenka</span>
    <span><i class="pill p-pn">PN</i> práce neschopnosti</span>
    <span><i class="pill p-x">X</i> nedostupný</span>
  </div>

  <div class="card todaybox" id="todaybox">
    <div class="card-h"><span>Dnes a zajtra</span><small id="todaynote"></small></div>
    <div class="tw">
      <table>
        <thead><tr><th class="emp">Pracovník</th><th id="thToday">Dnes</th><th id="thTomm">Zajtra</th></tr></thead>
        <tbody id="todaybody"><tr><td colspan="3" style="padding:14px;color:#6b7280">Načítavam…</td></tr></tbody>
      </table>
    </div>
  </div>

  <div class="card">
    <div class="card-h"><span id="title">Rozpis zmien</span><small>Klikni na svoje meno (len tvoje smeny) alebo na deň (kto pracuje)</small></div>
    <div class="mebar" id="mebar"></div>
    <div class="tw">
      <table>
        <thead><tr id="head"></tr></thead>
        <tbody id="body"><tr><td colspan="32" style="padding:24px;color:#6b7280">Načítavam rozpis…</td></tr></tbody>
      </table>
    </div>
    <div class="daybox" id="daybox"></div>
    <div class="foot">Sivé dni už uplynuli. Žlté záhlavie je dnešný deň.</div>
  </div>

  <div class="card chgbox" id="chgbox">
    <div class="card-h"><span>Nedávne zmeny v rozpise</span><small>zobrazujú sa 48 hodín</small></div>
    <ul id="chglist"></ul>
  </div>

  <div class="card rewards">
    <div class="card-h"><span>Odmeny: vnútorná obsluha</span><small>Kliknutím na meno uvidíš rozpis podľa kategórií</small></div>
    <div class="tw"><table><thead><tr><th class="emp">Pracovník</th><th>Poradie</th><th>Plnenie</th><th>Bonus celkom</th></tr></thead><tbody id="vno"></tbody></table></div>
  </div>

  <div class="card rewards">
    <div class="card-h"><span>Odmeny: vonkajšia obsluha</span><small>Kliknutím na meno uvidíš rozpis podľa kategórií</small></div>
    <div class="tw"><table><thead><tr><th class="emp">Pracovník</th><th>Poradie</th><th>Plnenie</th><th>Bonus celkom</th></tr></thead><tbody id="vo"></tbody></table></div>
  </div>
</div>

<script>
const DOW=["ne","po","ut","st","št","pi","so"];
const DOWFULL=["nedeľa","pondelok","utorok","streda","štvrtok","piatok","sobota"];
let me="",cache=null,todayData=null,selDay=null;
try{me=localStorage.getItem("me")||""}catch(e){}
const esc=s=>String(s).replace(/[&<>"']/g,c=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}[c]));

function pill(code){
  const c=String(code||"").toUpperCase();
  // dlhšie kódy musia byť skôr než kratšie (D-VO a DOV pred D)
  if(c.startsWith("D-VO"))return"p-dvo";
  if(c.startsWith("DOV"))return"p-dov";
  if(c.startsWith("PN"))return"p-pn";
  if(c.startsWith("POL"))return"p-pol";
  if(c.startsWith("D"))return"p-d";
  if(c.startsWith("N"))return"p-n";
  if(c.startsWith("X"))return"p-x";
  return"";
}
// skráti kód na samotný názov ("POL 18:00-22:30" -> "POL"), plný kód ostáva v názve bunky
function shortCode(code){
  const s=String(code||"").trim();
  const m=s.match(/^([A-Za-z]+(?:-[A-Za-z]+)?)/);
  return m?m[1].toUpperCase():s.toUpperCase();
}
function cell(code){
  if(!code)return"";
  return`<span class="pill ${pill(code)}" title="${esc(code)}">${esc(shortCode(code))}</span>`;
}
function isPast(y,m,d){
  const n=new Date();
  return new Date(y,m,d)<new Date(n.getFullYear(),n.getMonth(),n.getDate());
}
function isToday(y,m,d){
  const n=new Date();
  return y===n.getFullYear()&&m===n.getMonth()&&d===n.getDate();
}

async function showMonth(id){
  const body=document.getElementById("body");
  selDay=null;
  document.getElementById("daybox").style.display="none";
  try{
    const r=await fetch(`./data/${id}.json?v=${Date.now()}`);
    if(!r.ok)throw 0;
    const data=await r.json();
    document.getElementById("title").textContent="Rozpis zmien – "+data.label;
    if(data.updated){
      const u=new Date(data.updated);
      document.getElementById("upd").textContent="Aktualizované "+u.toLocaleString("sk-SK",{day:"numeric",month:"numeric",hour:"2-digit",minute:"2-digit"});
    }
    const cls=d=>{
      const w=new Date(data.y,data.m,d).getDay();
      return[(w===0||w===6)?"we":"",isPast(data.y,data.m,d)?"past":"",isToday(data.y,data.m,d)?"today":""].join(" ").trim();
    };
    document.getElementById("head").innerHTML='<th class="emp">Zamestnanec</th>'+
      data.days.map(d=>`<th class="day ${cls(d)}" onclick="showDay(${d})" title="Klikni: kto pracuje ${d}. ${data.m+1}."><span class="dn">${d}</span><span class="dw">${DOW[new Date(data.y,data.m,d).getDay()]}</span></th>`).join("");
    body.innerHTML=data.shifts.map(row=>
      `<tr data-n="${esc(row.n)}"><td class="emp" onclick="pick(this.parentNode.dataset.n)">${esc(row.n)}</td>`+
      data.days.map(d=>{
        const v=row.s[d]||"";
        const c=fresh(data,row.n,d);
        return`<td class="${(cls(d).replace("today","")+(c?" chg":"")).trim()}"${c?` title="Bolo: ${esc(c.was)}"`:""}>${cell(v)}</td>`;
      }).join("")+"</tr>").join("");
    rewards(data.vno,"vno",data);
    rewards(data.vo,"vo",data);
    cache=data;
    applyMe();
  }catch(e){
    body.innerHTML='<tr><td colspan="32" style="padding:24px;color:#dc2626">Dáta pre vybraný mesiac sa nepodarilo načítať.</td></tr>';
  }
}

/* ---------- klik na deň: kto pracuje ---------- */
function showDay(d){
  const data=cache;
  if(!data)return;
  selDay=(selDay===d)?null:d;
  document.querySelectorAll("th.day").forEach(th=>th.classList.remove("sel"));
  const box=document.getElementById("daybox");
  if(selDay===null){box.style.display="none";box.innerHTML="";return;}
  const th=[...document.querySelectorAll("th.day")].find(x=>x.querySelector(".dn").textContent==String(d));
  if(th)th.classList.add("sel");
  const w=new Date(data.y,data.m,d).getDay();
  const rows=data.shifts.map(r=>({n:r.n,v:r.s[d]||""})).filter(x=>x.v);
  const work=rows.filter(x=>!isOff(x.v));
  const off=rows.filter(x=>isOff(x.v));
  let h=`<div class="dh">${d}. ${data.m+1}. ${data.y} – ${DOWFULL[w]}</div>`;
  h+='<div class="dl">';
  h+=work.length
      ? work.map(x=>`<span class="one"><b>${esc(x.n)}</b>${cell(x.v)}</span>`).join("")
      : '<span class="none">V tento deň nikto nepracuje.</span>';
  h+="</div>";
  if(off.length){
    h+='<div class="dl" style="margin-top:8px">'+off.map(x=>
      `<span class="one off"><b>${esc(x.n)}</b>${cell(x.v)}</span>`).join("")+"</div>";
  }
  box.innerHTML=h;
  box.style.display="block";
}
function isOff(code){
  const c=String(code||"").toUpperCase();
  return c.startsWith("X")||c.startsWith("DOV")||c.startsWith("PN");
}

/* ---------- panel Dnes a zajtra ---------- */
async function loadToday(){
  const n=new Date();
  const cur=`${n.getFullYear()}-${String(n.getMonth()+1).padStart(2,"0")}`;
  try{
    const r=await fetch(`./data/${cur}.json?v=${Date.now()}`);
    if(!r.ok)throw 0;
    todayData=await r.json();
  }catch(e){todayData=null;}
  renderToday();
}
function renderToday(){
  const tb=document.getElementById("todaybody");
  const note=document.getElementById("todaynote");
  if(!todayData){tb.innerHTML='<tr><td colspan="3" style="padding:14px;color:#6b7280">Dnešná smenovka sa nepodarilo načítať.</td></tr>';note.textContent="";return;}
  const data=todayData, y=data.y, m=data.m;
  const n=new Date();
  const d1=n.getDate(), d2=d1+1;
  const label=(d)=>`${DOWFULL[new Date(y,m,d).getDay()]} ${d}. ${m+1}.`;
  document.getElementById("thToday").innerHTML=`Dnes<br><span class="wd">${label(d1)}</span>`;
  document.getElementById("thTomm").innerHTML=`Zajtra<br><span class="wd">${label(d2)}</span>`;
  note.textContent=data.label;
  const has=d=>data.days.includes(d);
  const names=data.shifts.map(r=>r.n);
  let rows="";
  names.forEach(nm=>{
    const r=data.shifts.find(x=>x.n===nm);
    const v1=has(d1)?(r.s[d1]||""):"";
    const v2=has(d2)?(r.s[d2]||""):"";
    if(!v1&&!v2)return;
    rows+=`<tr><td class="emp">${esc(nm)}</td><td>${cell(v1)||'<span class="none">–</span>'}</td><td>${cell(v2)||'<span class="none">–</span>'}</td></tr>`;
  });
  tb.innerHTML=rows||'<tr><td colspan="3" style="padding:14px;color:#6b7280">V dnešnom mesiaci nie sú zaznamenané žiadne smeny.</td></tr>';
}

const HOURS48=48*3600*1000;
function fresh(data,name,d){
  const c=data.changes&&data.changes[name]&&data.changes[name][d];
  if(!c||Date.now()-new Date(c.at)>HOURS48)return null;
  return isPast(data.y,data.m,d)?null:c;
}
function changes(data){
  const items=[];
  data.shifts.forEach(row=>data.days.forEach(d=>{
    if(me&&row.n!==me)return;
    const c=fresh(data,row.n,d);
    if(c)items.push(`<li><span class="dot"></span><strong>${esc(row.n)}</strong>, ${d}. ${data.m+1}.: bolo ${cell(c.was)} teraz ${cell(row.s[d]||"-")}</li>`);
  }));
  document.getElementById("chglist").innerHTML=items.join("");
  document.getElementById("chgbox").style.display=items.length?"block":"none";
}

function pick(name){
  me=(me===name)?"":name;
  try{me?localStorage.setItem("me",me):localStorage.removeItem("me")}catch(e){}
  applyMe();
}
function applyMe(){
  const rows=document.querySelectorAll("#body tr[data-n]");
  if(me&&![...rows].some(r=>r.dataset.n===me))me="";   // meno v tomto mesiaci nie je
  rows.forEach(r=>r.style.display=(me&&r.dataset.n!==me)?"none":"");
  const bar=document.getElementById("mebar");
  if(me){
    bar.style.display="flex";
    bar.innerHTML=`<span>Zobrazené sú iba smeny: <strong>${esc(me)}</strong></span><button onclick="pick(me)">Zobraziť všetkých</button>`;
  }else bar.style.display="none";
  if(cache)changes(cache);
}

const norm=s=>String(s).normalize("NFD").replace(/[\u0300-\u036f]/g,"").toLowerCase();
const baseName=s=>norm(s).replace(/\(.*?\)/g," ").replace(/\s+/g," ").trim();
/* Plány sú podľa typu smeny (Excel, list PLANY):
   vno = vnútorná obsluha (D = vnútorná denná, N = vnútorná nočná)
   vo  = vonkajšia obsluha (D-VO = denná vonku)                        */
const REWTYPES={vno:["D","N"],vo:["D-VO"]};
const TYPELBL={D:"denná smena",N:"nočná smena","D-VO":"denná smena vonku"};
// zostávajúce smeny po type (od dneška); null = v rozpise sa nenašiel
function shiftsByType(data,id,name){
  if(!data||!data.shifts)return null;
  const b=baseName(name);
  let rows=data.shifts.filter(r=>baseName(r.n)===b);
  if(rows.length>1){
    const isVo=r=>/\(vo\)/.test(norm(r.n));
    const pickRows=id==="vo"?rows.filter(isVo):rows.filter(r=>!isVo(r));
    if(pickRows.length)rows=pickRows;
  }
  if(!rows.length)return null;
  const cnt={D:0,N:0,"D-VO":0,POL:0};
  data.days.forEach(d=>{
    if(isPast(data.y,data.m,d))return;
    rows.forEach(r=>{
      const v=(r.s[d]||"").toUpperCase();
      if(!v)return;
      if(v.startsWith("D-VO")){cnt["D-VO"]++;return;}                                  // denná vonku
      if(v.startsWith("X")||v.startsWith("DOV")||v.startsWith("PN"))return;          // nepracuje
      if(v.startsWith("POL")){cnt.POL++;return;}                                      // krátka polnočná smena (bez plánu)
      if(v.startsWith("N")){cnt.N++;return;}
      if(v.startsWith("D")){cnt.D++;}
    });
  });
  return cnt;
}
// denný plán tovaru pre daný typ smeny (z Excelu, list PLANY)
function planOf(data,type,cat){
  const P=(data&&data.plans&&data.plans[type])||{};
  const v=P[norm(cat).toUpperCase()];
  return (v&&typeof v.p==="number")?v:null;
}
/* Koľko ešte treba predať – počíta sa podľa typov zostávajúcich smien:
   plán zostávajúcich smien = denný plán (PLANY) × počet smien daného typu,
   deficit (mesačný plán − skutočnosť) sa rozdelí medzi typy úmerne ich plánu. */
function gap2(data,id,c,cnt){
  const plan=c[1],act=c[2];
  if(!plan)return"";
  const types=REWTYPES[id]||[];
  const left=types.reduce((s,t)=>s+(cnt?cnt[t]:0),0);
  const pol=cnt?cnt.POL:0;
  const g90=Math.max(0,plan*0.9-act),g100=Math.max(0,plan-act);
  const c90=Math.ceil(g90-1e-9),c100=Math.ceil(g100-1e-9);
  const rows=[];let sum=0;
  types.forEach(t=>{
    const n=cnt?cnt[t]:0;
    if(!n)return;
    const p=planOf(data,t,c[0]);
    if(!p)return;                       // tento tovar nemá plán pre daný typ smeny
    const v=p.p*n;sum+=v;
    rows.push({t:t,n:n,v:v,u:p.u||""});
  });
  if(!rows.length){                       // zostávajúce smeny nemajú pre tento tovar plán
    if(g100<=0&&g90<=0)return'<span class="ok">✓ Plán 100 % splnený</span>';
    let pre="";
    if(cnt===null)pre="";
    else if(left===0)pre=pol?`Zostávajú len krátke polnočné smeny (${pol}× POL), tie nemajú vlastný plán. `
      :`Už nemá žiadne naplánované ${id==="vo"?"vonkajšie":"vnútorné"} smeny. `;
    return`${pre}Do 90 % chýba <strong>${c90}</strong>, do 100 % chýba <strong>${c100}</strong>.`;
  }
  const per=(r,g)=>Math.ceil((r.v+g*r.v/sum)/r.n-1e-9);   // koľko na jednu smenu daného typu
  const line=g=>rows.map(r=>`${TYPELBL[r.t]} (${r.n}×) <strong>${per(r,g)}${r.u?" "+r.u:""}</strong>`).join(" · ");
  if(g100>0){
    let h=g90>0?`do 90 %: ${line(g90)}<br>`:"";
    return h+`do 100 %: ${line(g100)}`;
  }
  return`<span class="ok">✓ Plán 100 % splnený</span><br>plán na zostávajúce smeny: ${line(0)}`;
}

function rewards(list,id,data){
  const tb=document.getElementById(id);
  if(!list||!list.length){tb.innerHTML='<tr><td colspan="4" style="padding:14px;color:#6b7280">Žiadne záznamy odmien.</td></tr>';return;}
  tb.innerHTML=list.map((it,i)=>{
    const sid=`${id}-${i}`, has=it.c&&it.c.length, cnt=shiftsByType(data,id,it.n);
    const mine=!!(me&&baseName(it.n)===baseName(me));
    let h=`<tr class="rw${mine?" mine":""}" onclick="tog('${sid}')"><td class="emp">${esc(it.n)}${has?' <span style="font-size:10px;color:#9aa1ab">▼</span>':""}</td>
      <td>${it.rank!=null?it.rank+".":"-"}</td>
      <td>${it.pl?(it.pl*100).toFixed(1)+" %":"-"}</td>
      <td><strong>${it.tot!=null?it.tot.toFixed(2)+" €":"-"}</strong></td></tr>`;
    if(has)h+=`<tr id="${sid}" class="dt${mine?" mine":""}"><td colspan="4">`+it.c.map(c=>
      `<span class="chip"><strong>${esc(c[0])}:</strong> plán ${c[1]} · skutočnosť ${c[2]} (${(c[3]*100).toFixed(0)} %) · <strong>${c[4].toFixed(2)} €</strong><br><span class="gap">${gap2(data,id,c,cnt)}</span></span>`).join("")+"</td></tr>";
    return h;
  }).join("");
}
function tog(id){const e=document.getElementById(id);if(e)e.style.display=e.style.display==="table-row"?"none":"table-row";}

async function init(){
  const nav=document.getElementById("month-list");
  loadToday();
  try{
    const r=await fetch(`./data/index.json?v=${Date.now()}`);
    if(!r.ok)throw 0;
    const idx=await r.json();
    if(!idx.months||!idx.months.length){document.getElementById("body").innerHTML='<tr><td colspan="32" style="padding:24px;color:#6b7280">Žiadne dostupné smenovky.</td></tr>';return;}
    const n=new Date(), cur=`${n.getFullYear()}-${String(n.getMonth()+1).padStart(2,"0")}`;
    const pick=idx.months.find(m=>m.id===cur)||idx.months[0];
    idx.months.forEach(m=>{
      const b=document.createElement("button");
      b.className="month-btn"+(m.id===pick.id?" active":"");
      b.textContent=m.label;
      b.onclick=()=>{nav.querySelectorAll(".month-btn").forEach(x=>x.classList.remove("active"));b.classList.add("active");showMonth(m.id);};
      nav.appendChild(b);
    });
    showMonth(pick.id);
  }catch(e){
    document.getElementById("body").innerHTML='<tr><td colspan="32" style="padding:24px;color:#dc2626">Chyba načítania zoznamu mesiacov.</td></tr>';
  }
}
init();
</script>
</body>
</html>
```
