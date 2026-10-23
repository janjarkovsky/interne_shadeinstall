# ShadeInstall

Desktopová aplikace ShadeInstall v CustomTkinteru s modulem přípravy cenové
nabídky. Sidebar obsahuje jen tento modul; profil a další moduly jsou odebrané.
Logo je zobrazené vedle názvu ShadeInstall. Zadané rozměry aplikace použije
k výpočtu plochy v m² a zapíše je do nového
listu `Výpočty ShadeInstall` v kopii vybraného sešitu `.xlsx` nebo `.xlsm`.
Původní soubor zůstává nezměněný a stávající listy se vzorci se nepřepisují.

Cena v Kč zatím není vypočítávána: k tomu jsou potřeba sazby a konkrétní
struktura ceníku.

## Vytvoření EXE ve Windows

V prostředí s nainstalovaným Pythonem spusťte v této složce:

```powershell
py -m pip install customtkinter openpyxl pyinstaller
py -m PyInstaller --noconsole --onefile --name ShadeInstall --collect-all customtkinter --icon=logo.ico --add-data=logo.ico:. --add-data=logo_mark.png:. main.pyw
```

Výsledný soubor `ShadeInstall.exe` bude ve složce `dist`. Cílový uživatel k jeho spuštění
nepotřebuje samostatně instalovat Python ani CustomTkinter.