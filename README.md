# ShadeInstall

Minimální bezokenní GUI základ v CustomTkinteru. Při spuštění zobrazí bílé
okno 800 x 600 px. Budoucí výpočet cen a práci s Excel ceníkem lze doplnit
do samostatné logiky; při ukládání upravujte jen určené buňky, aby zůstaly
zachované ostatní vzorce.

## Vytvoření EXE ve Windows

V prostředí s nainstalovaným Pythonem spusťte v této složce:

```powershell
py -m pip install customtkinter pyinstaller
py -m PyInstaller --noconsole --onefile --collect-all customtkinter main.pyw
```

Výsledný soubor bude ve složce `dist`. Cílový uživatel k jeho spuštění
nepotřebuje samostatně instalovat Python ani CustomTkinter.