# ShadeInstall

Desktopová aplikace ShadeInstall připravuje nabídky pro Zetta 90 s barvou Standard
nebo RAL a pohonem Motorický nebo Manuální. Podporuje až 14 položek stejného typu,
vyhledá cenu podle rozměrů a vyplní kopii odpovídajícího vzoru. Původní vzory
zůstávají nezměněné.

## Vytvoření EXE ve Windows

V prostředí s nainstalovaným Pythonem spusťte v této složce:

```powershell
py -m pip install customtkinter openpyxl pyinstaller
py -m PyInstaller --noconsole --onefile --name ShadeInstall --collect-all customtkinter --icon=logo.ico --add-data=logo.ico:. --add-data=logo_mark.png:. --add-data=data/ceny_zetta90.xlsx:data --add-data=data/vzor_cenik_zetta90_standard_motor.xlsx:data --add-data=data/vzor_cenik_Zetta90_standard_manual.xlsx:data --add-data=data/vzor_cenik_Zetta90_ral_motor.xlsx:data --add-data=data/vzor_cenik_Zetta90_ral_manual.xlsx:data main.pyw
```

Výsledný soubor `ShadeInstall.exe` bude ve složce `dist`. Cílový uživatel k jeho spuštění
nepotřebuje samostatně instalovat Python ani CustomTkinter.