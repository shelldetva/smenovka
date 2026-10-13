# Smenovka na web: návod krok za krokom

Všetko je zadarmo. Jednorazové nastavenie trvá približne hodinu, potom už nerobíš nič iné, len ukladáš Excel.

## Ako to funguje
1. Uložíš mesačný Excel do priečinka v počítači (napr. `Master_Smenovka_2026_11.xlsx`).
2. Program `sync.py` (beží na pozadí) si to všimne, vytiahne smenovku a odmeny a nahrá ich na web.
3. Zamestnanci otvoria stály odkaz a vidia aktuálnu smenovku. Stránka sa sama obnovuje každých 5 minút.

## Jednorazové nastavenie

### 1. Python
- Stiahni z python.org a nainštaluj. **Pri inštalácii zaškrtni „Add Python to PATH".**
- Otvor Príkazový riadok (Win + R, napíš `cmd`) a spusti: `pip install openpyxl`

### 2. Git a GitHub
- Nainštaluj Git z git-scm.com (stačia predvolené voľby).
- Založ si bezplatný účet na github.com.
- Na GitHube vytvor nový repozitár s názvom `smenovka` (tlačidlo „New"). Musí byť **Public** (bezplatné stránky to vyžadujú).

### 3. Priečinok s programom
- Rozbaľ tento balík napr. do `C:\SmenovkaWeb`.
- Otvor v ňom príkazový riadok (do adresného riadka priečinka napíš `cmd` a Enter) a spusti postupne (za `MENO` dosaď svoje meno na GitHube):
```
git init
git branch -M main
git remote add origin https://github.com/MENO/smenovka.git
git add -A
git commit -m "start"
git push -u origin main
```
Pri prvom `push` sa otvorí okno na prihlásenie do GitHubu, prihlás sa.

### 4. Zapnutie webu
- Na GitHube: repozitár `smenovka` → **Settings** → **Pages**.
- V časti „Branch" vyber `main` a priečinok `/docs`, potom **Save**.
- Po minúte-dvoch bude web na adrese `https://MENO.github.io/smenovka/`. Túto adresu pošleš zamestnancom.

### 5. Nastavenie programu
- Otvor `sync.py` v Poznámkovom bloku a hore zmeň riadok `EXCEL_FOLDER` na priečinok, kde ukladáš Excely, napr. `Path(r"C:\Smenovka")`.
- Excely musia mať názov v tvare `Master_Smenovka_ROK_MESIAC.xlsx`, napr. `Master_Smenovka_2026_11.xlsx`.

### 6. Prvé spustenie
- Spusti `spusti.bat` (dvojklik). V okne uvidíš „aktualizované" a „web nahraný".
- Po minúte-dvoch otvor adresu webu a skontroluj to.

### 7. Aby to bežalo samo
- Stlač Win + R, napíš `shell:startup` a Enter.
- Do otvoreného priečinka vlož odkaz na `spusti.bat` (pravý klik → Vytvoriť odkaz). Program sa potom spustí s počítačom. Okno nezatváraj, stačí ho minimalizovať.

## Bežná prevádzka
- **Zmena smien alebo odmien:** uprav Excel, ulož. Zmeny sa na web nahrávajú najviac raz za hodinu (kvôli limitu bezplatného hostingu), takže sa prejavia najneskôr za hodinu a pár minút. Interval zmeníš v `sync.py` hodnotou `PUSH_EVERY_MINUTES`.
- **Nový mesiac:** ulož nový súbor s názvom podľa vzoru. Na webe pribudne nová karta.

## Dobré vedieť
- Počítač s programom musí byť zapnutý a pripojený na internet, keď ukladáš Excel.
- Web je verejný: kto pozná adresu, vidí aj odmeny. Adresu posielaj len zamestnancom.
- Program číta konkrétne miesta v zošite (listy `MASTER_SMENOVKA` a `SUMAR`). Ak rozloženie zmeníš (presun stĺpcov, nové riadky nad tabuľkami), treba program upraviť.
- Ak sa niečo pokazí, v okne programu uvidíš chybovú správu. Pošli mi ju a vyriešime to.
