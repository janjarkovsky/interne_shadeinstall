# ShadeInstall

Desktopová aplikace ShadeInstall připravuje cenové nabídky pro žaluzie Zetta 90.
Podporuje až 14 položek stejného typu v jedné nabídce, vyhledá cenu podle
rozměrů a vyplní kopii vzoru `data/cenik_vzor_zetta90.xlsx`. Původní vzor zůstává
nezměněný.

## Vytvoření EXE ve Windows

V prostředí s nainstalovaným Pythonem spusťte v této složce:

```powershell
py -m pip install customtkinter openpyxl pyinstaller
py -m PyInstaller --noconsole --onefile --name ShadeInstall --collect-all customtkinter --icon=logo.ico --add-data=logo.ico:. --add-data=logo_mark.png:. --add-data=data/ceny_zetta90.xlsx:data --add-data=data/cenik_vzor_zetta90.xlsx:data main.pyw
```

Výsledný soubor `ShadeInstall.exe` bude ve složce `dist`. Cílový uživatel k jeho spuštění
nepotřebuje samostatně instalovat Python ani CustomTkinter.