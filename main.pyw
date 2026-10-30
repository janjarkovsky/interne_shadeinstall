"""ShadeInstall - Interní Aplikace."""

import re
import sys
from datetime import datetime
from decimal import Decimal, InvalidOperation
from pathlib import Path
from tkinter import filedialog, messagebox

import customtkinter as ctk
from openpyxl import load_workbook
from PIL import Image


def resource_path(filename: str) -> Path:
    """Najde soubor vedle skriptu i uvnitř PyInstaller EXE."""
    base_path = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))
    return base_path / filename


def parse_dimension(value: str, name: str) -> Decimal:
    try:
        dimension = Decimal(value.strip().replace(",", "."))
    except InvalidOperation as error:
        raise ValueError(f"Zadejte platnou číselnou hodnotu pro {name}.") from error
    if not dimension.is_finite() or dimension <= 0:
        raise ValueError(f"{name.capitalize()} musí být kladné číslo.")
    return dimension


def parse_count(value: str) -> int:
    try:
        count = int(value.strip())
    except ValueError as error:
        raise ValueError("Počet musí být kladné celé číslo.") from error
    if count <= 0:
        raise ValueError("Počet musí být kladné celé číslo.")
    return count


def calculate_holder_count(width_cm: Decimal) -> int:
    if width_cm <= 160:
        return 2
    if width_cm <= 240:
        return 3
    if width_cm <= 320:
        return 4
    if width_cm <= 420:
        return 5
    return 6


TEMPLATE_FILE = "data/cenik_vzor.xlsx"
PRICE_FILES = {"Zetta 90": "data/ceny_zetta90.xlsx"}
TEMPLATE_CELLS = {
    "width_cm": "E23",
    "height_cm": "F23",
    "holder_count": "C23",
    "count": "G23",
    "price": "I23",
    "quote_number": "K9",
}


def lookup_price(blind_type: str, width_mm: Decimal, height_mm: Decimal) -> int:
    price_file = PRICE_FILES.get(blind_type)
    if price_file is None:
        raise ValueError(f"Pro žaluzii {blind_type} není k dispozici ceník.")

    price_path = resource_path(price_file)
    if not price_path.exists():
        raise FileNotFoundError(f"Chybí cenová tabulka {price_file}.")

    workbook = load_workbook(price_path, read_only=True, data_only=True)
    try:
        worksheet = workbook.active
        widths = [worksheet.cell(row=1, column=column).value
                  for column in range(2, worksheet.max_column + 1)]
        heights = [worksheet.cell(row=row, column=1).value
                   for row in range(2, worksheet.max_row + 1)]
        selected_width = next(
            (value for value in widths if Decimal(str(value)) >= width_mm), None
        )
        selected_height = next(
            (value for value in heights if Decimal(str(value)) >= height_mm), None
        )
        if selected_width is None or selected_height is None:
            raise ValueError("Zadané rozměry přesahují rozsah cenové tabulky.")

        column = widths.index(selected_width) + 2
        row = heights.index(selected_height) + 2
        price = worksheet.cell(row=row, column=column).value
        if not isinstance(price, (int, float)):
            raise ValueError("Pro zadané rozměry není v cenové tabulce cena.")
        return int(price)
    finally:
        workbook.close()


def create_pricelist(
    output_path: Path,
    width_cm: Decimal,
    height_cm: Decimal,
    holder_count: int,
    count: int,
    quote_number: str,
    price: int,
) -> None:
    """Vyplní kopii šablony zadanými parametry a uloží ji jako nový soubor."""
    template_path = resource_path(TEMPLATE_FILE)
    if not template_path.exists():
        raise FileNotFoundError(f"Chybí šablona {TEMPLATE_FILE}.")
    workbook = load_workbook(template_path)
    try:
        worksheet = workbook.active
        values = {
            "width_cm": float(width_cm),
            "height_cm": float(height_cm),
            "holder_count": holder_count,
            "count": count,
            "quote_number": quote_number,
            "price": price,
        }
        for key, cell in TEMPLATE_CELLS.items():
            worksheet[cell] = values[key]
        workbook.save(output_path)
    finally:
        workbook.close()


BLIND_TYPES = ["Zetta 90"]


class ShadeInstallApp(ctk.CTk):
    def __init__(self) -> None:
        super().__init__()
        self.title("ShadeInstall - Interní Aplikace")
        self.geometry("1365x910")
        self.minsize(850, 600)
        self.configure(fg_color="#f8f9fa")
        self.iconbitmap(str(resource_path("logo.ico")))
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        self.sidebar_frame = ctk.CTkFrame(
            self, width=260, corner_radius=0, fg_color="#ffffff"
        )
        self.sidebar_frame.grid(row=0, column=0, sticky="nsew")
        self.sidebar_frame.grid_propagate(False)
        self.sidebar_frame.grid_columnconfigure(0, weight=1)

        brand_frame = ctk.CTkFrame(self.sidebar_frame, fg_color="transparent")
        brand_frame.grid(row=0, column=0, padx=20, pady=(25, 20), sticky="w")
        self.brand_image = ctk.CTkImage(
            light_image=Image.open(resource_path("logo_mark.png")),
            size=(42, 42),
        )
        ctk.CTkLabel(brand_frame, text="", image=self.brand_image).pack(side="left")
        ctk.CTkLabel(
            brand_frame,
            text="ShadeInstall.cz",
            font=("Segoe UI", 16, "bold"),
            text_color="#0d6efd",
        ).pack(side="left", padx=(10, 0))
        ctk.CTkLabel(
            self.sidebar_frame,
            text="MODULY",
            font=("Segoe UI", 10, "bold"),
            text_color="#94a3b8",
        ).grid(row=1, column=0, padx=20, pady=(10, 5), sticky="w")

        ctk.CTkButton(
            self.sidebar_frame,
            text="Příprava cenové nabídky",
            fg_color="#0d6efd",
            hover_color="#0b5ed7",
            text_color="#ffffff",
            corner_radius=8,
            anchor="w",
            font=("Segoe UI", 13, "bold"),
            height=42,
        ).grid(row=2, column=0, padx=12, pady=4, sticky="ew")

        self.main_container = ctk.CTkFrame(
            self, fg_color="#f8f9fa", corner_radius=0
        )
        self.main_container.grid(row=0, column=1, sticky="nsew")
        self.main_container.grid_columnconfigure(0, weight=1)
        self.main_container.grid_rowconfigure(1, weight=1)

        self.top_bar = ctk.CTkFrame(
            self.main_container, height=65, fg_color="#ffffff", corner_radius=0
        )
        self.top_bar.grid(row=0, column=0, sticky="ew")
        self.top_bar.grid_propagate(False)
        self.top_bar.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(
            self.top_bar,
            text="Příprava cenové nabídky",
            font=("Segoe UI", 18, "bold"),
            text_color="#1e293b",
        ).grid(row=0, column=0, padx=25, pady=15, sticky="w")

        content = ctk.CTkFrame(
            self.main_container, fg_color="#ffffff", corner_radius=12
        )
        content.grid(row=1, column=0, padx=25, pady=25, sticky="nsew")

        self.entry_customer_name = self._add_text_field(
            content, "Jméno zákazníka:"
        )
        ctk.CTkLabel(
            content,
            text="Výběr žaluzie:",
            font=("Segoe UI", 13),
            text_color="#475569",
        ).pack(anchor="w", padx=30, pady=(5, 2))
        self.blind_type = ctk.CTkOptionMenu(
            content,
            values=BLIND_TYPES,
            width=320,
            height=38,
            corner_radius=8,
            fg_color="#f8f9fa",
            button_color="#cbd5e1",
            button_hover_color="#94a3b8",
            text_color="#1e293b",
        )
        self.blind_type.pack(anchor="w", padx=30, pady=(0, 12))

        self.entry_width = self._add_text_field(content, "Šířka (cm):")
        self.entry_height = self._add_text_field(content, "Výška (cm):")
        self.entry_count = self._add_text_field(content, "Počet:")
        self.entry_quote_number = self._add_text_field(
            content, "Číslo cenové nabídky:"
        )
        self.price_result = ctk.CTkLabel(
            content,
            text="",
            font=("Segoe UI", 16, "bold"),
            text_color="#16803c",
        )
        self.price_result.pack(anchor="w", padx=30, pady=(0, 8))
        ctk.CTkButton(
            content,
            text="Vytvořit ceník",
            fg_color="#0d6efd",
            hover_color="#0b5ed7",
            text_color="#ffffff",
            corner_radius=8,
            font=("Segoe UI", 13, "bold"),
            width=320,
            height=42,
            command=self.calculate_and_write,
        ).pack(anchor="w", padx=30, pady=10)

    @staticmethod
    def _add_text_field(parent: ctk.CTkFrame, label: str) -> ctk.CTkEntry:
        ctk.CTkLabel(
            parent, text=label, font=("Segoe UI", 13), text_color="#475569"
        ).pack(anchor="w", padx=30, pady=(5, 2))
        entry = ctk.CTkEntry(
            parent,
            width=320,
            height=38,
            corner_radius=8,
            fg_color="#f8f9fa",
            border_color="#cbd5e1",
            text_color="#1e293b",
        )
        entry.pack(anchor="w", padx=30, pady=(0, 12))
        return entry

    def calculate_and_write(self) -> None:
        customer_name = self.entry_customer_name.get().strip()
        if not customer_name:
            messagebox.showerror(
                "Chybí jméno zákazníka",
                "Zadejte jméno zákazníka.",
                parent=self,
            )
            return
        customer_name = re.sub(
            r'[<>:"/\\|?*\x00-\x1f]', "_", customer_name
        ).strip(" .")
        if not customer_name:
            messagebox.showerror(
                "Neplatné jméno zákazníka",
                "Zadejte jméno obsahující znaky použitelné v názvu souboru.",
                parent=self,
            )
            return

        quote_number = self.entry_quote_number.get().strip()
        if not quote_number:
            messagebox.showerror(
                "Chybí číslo cenové nabídky",
                "Zadejte číslo cenové nabídky.",
                parent=self,
            )
            return
        quote_number = f"{quote_number}-{datetime.now():%y}"

        try:
            width_cm = parse_dimension(self.entry_width.get(), "šířku")
            height_cm = parse_dimension(self.entry_height.get(), "výšku")
            count = parse_count(self.entry_count.get())
        except ValueError as error:
            messagebox.showerror("Neplatné údaje", str(error), parent=self)
            return

        holder_count = calculate_holder_count(width_cm)
        width_mm = width_cm * 10
        height_mm = height_cm * 10

        try:
            price = lookup_price(self.blind_type.get(), width_mm, height_mm)
        except (FileNotFoundError, ValueError) as error:
            self.price_result.configure(text="")
            messagebox.showerror("Cenu se nepodařilo vypočítat", str(error), parent=self)
            return
        self.price_result.configure(text=f"Výsledná cena: {price:,} Kč".replace(",", " "))

        selected_file = filedialog.asksaveasfilename(
            title="Uložit nový ceník",
            defaultextension=".xlsx",
            initialfile=(
                f"Ceník ({customer_name}) ({self.blind_type.get()}).xlsx"
            ),
            filetypes=(("Excel sešit", "*.xlsx"),),
            parent=self,
        )
        if not selected_file:
            return
        output_path = Path(selected_file)
        try:
            create_pricelist(
                output_path,
                width_cm,
                height_cm,
                holder_count,
                count,
                quote_number,
                price,
            )
        except Exception as error:
            messagebox.showerror(
                "Vytvoření ceníku se nezdařilo", str(error), parent=self
            )
            return

        messagebox.showinfo(
            "Ceník vytvořen", f"Nový ceník:\n{output_path}", parent=self
        )


def main() -> None:
    ctk.set_appearance_mode("light")
    ctk.set_default_color_theme("blue")
    ShadeInstallApp().mainloop()


if __name__ == "__main__":
    main()