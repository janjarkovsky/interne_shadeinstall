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


TEMPLATE_FILE = "data/cenik_vzor_zetta90.xlsx"
PRICE_FILES = {"Zetta 90": "data/ceny_zetta90.xlsx"}
ITEM_START_ROW = 23
MAX_BLIND_ITEMS = 14


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
    items: list[dict],
    quote_number: str,
) -> None:
    """Vyplní kopii šablony zadanými parametry a uloží ji jako nový soubor."""
    if not 1 <= len(items) <= MAX_BLIND_ITEMS:
        raise ValueError(f"Ceník musí obsahovat 1 až {MAX_BLIND_ITEMS} žaluzií.")
    template_path = resource_path(TEMPLATE_FILE)
    if not template_path.exists():
        raise FileNotFoundError(f"Chybí šablona {TEMPLATE_FILE}.")
    workbook = load_workbook(template_path)
    try:
        worksheet = workbook.active
        worksheet["K9"] = quote_number
        for item_index, item in enumerate(items):
            row = ITEM_START_ROW + item_index
            worksheet[f"B{row}"] = item["holder_count"]
            worksheet[f"E{row}"] = float(item["width_cm"])
            worksheet[f"F{row}"] = float(item["height_cm"])
            worksheet[f"G{row}"] = item["count"]
            worksheet[f"I{row}"] = item["price"]
        workbook.calculation.fullCalcOnLoad = True
        workbook.calculation.forceFullCalc = True
        workbook.calculation.calcMode = "auto"
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

        content = ctk.CTkScrollableFrame(
            self.main_container,
            fg_color="#ffffff",
            corner_radius=12,
            scrollbar_button_color="#cbd5e1",
            scrollbar_button_hover_color="#94a3b8",
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

        self.entry_quote_number = self._add_text_field(
            content, "Číslo cenové nabídky:"
        )
        self.blind_items_frame = ctk.CTkFrame(content, fg_color="transparent")
        self.blind_items_frame.pack(fill="x", padx=30, pady=(8, 8))
        for column, label in enumerate(
            ("Žaluzie", "Šířka (cm)", "Výška (cm)", "Počet")
        ):
            ctk.CTkLabel(
                self.blind_items_frame,
                text=label,
                font=("Segoe UI", 12, "bold"),
                text_color="#475569",
            ).grid(row=0, column=column, padx=5, pady=(0, 4), sticky="w")
        self.blind_items_frame.grid_columnconfigure(0, minsize=90)
        for column in range(1, 4):
            self.blind_items_frame.grid_columnconfigure(column, weight=1)

        self.blind_items = []
        self.add_blind_button = ctk.CTkButton(
            content,
            text="Přidat další žaluzii",
            fg_color="#16803c",
            hover_color="#126c32",
            text_color="#ffffff",
            corner_radius=8,
            font=("Segoe UI", 12, "bold"),
            width=260,
            height=34,
            command=lambda: self.add_blind_item(focus_new=True),
        )
        self.add_blind_button.pack(anchor="w", padx=30, pady=(0, 12))
        self.add_blind_item(focus_new=False)
        self.price_result = ctk.CTkLabel(
            content,
            text="",
            font=("Segoe UI", 16, "bold"),
            text_color="#16803c",
        )
        self.price_result.pack(anchor="w", padx=30, pady=(0, 8))
        self.create_button = ctk.CTkButton(
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
        )
        self.create_button.pack(anchor="w", padx=30, pady=10)

        self._bind_enter(self.entry_customer_name, self.blind_type)
        self._bind_enter(self.blind_type, self.entry_quote_number)
        self._bind_enter(self.entry_quote_number, self.blind_items[0]["width"])
        self.add_blind_button.bind(
            "<Return>", self._handle_add_blind_enter, add="+"
        )
        self.create_button.bind(
            "<Return>", self._handle_create_enter, add="+"
        )

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

    @staticmethod
    def _create_item_entry(parent: ctk.CTkFrame, width: int) -> ctk.CTkEntry:
        return ctk.CTkEntry(
            parent,
            width=width,
            height=36,
            corner_radius=8,
            fg_color="#f8f9fa",
            border_color="#cbd5e1",
            text_color="#1e293b",
        )

    @staticmethod
    def _move_focus(_event, target) -> str:
        target.focus_set()
        return "break"

    def _bind_enter(self, widget, target) -> None:
        widget.bind(
            "<Return>",
            lambda event, next_widget=target: self._move_focus(
                event, next_widget
            ),
            add="+",
        )

    def _focus_after_item(self, event, item_index: int) -> str:
        next_index = item_index + 1
        if next_index < len(self.blind_items):
            target = self.blind_items[next_index]["width"]
        elif len(self.blind_items) < MAX_BLIND_ITEMS:
            target = self.add_blind_button
        else:
            target = self.create_button
        return self._move_focus(event, target)

    def _handle_add_blind_enter(self, _event) -> str:
        self.add_blind_item(focus_new=True)
        return "break"

    def _handle_create_enter(self, _event) -> str:
        self.calculate_and_write()
        return "break"

    def add_blind_item(self, focus_new: bool = True) -> None:
        if len(self.blind_items) >= MAX_BLIND_ITEMS:
            return

        item_number = len(self.blind_items) + 1
        row_frame = ctk.CTkFrame(self.blind_items_frame, fg_color="transparent")
        row_frame.grid(
            row=item_number,
            column=0,
            columnspan=4,
            padx=0,
            pady=3,
            sticky="ew",
        )
        row_frame.grid_columnconfigure(0, minsize=90)
        for column in range(1, 4):
            row_frame.grid_columnconfigure(column, weight=1)

        ctk.CTkLabel(
            row_frame,
            text=str(item_number),
            font=("Segoe UI", 13),
            text_color="#475569",
        ).grid(row=0, column=0, padx=5, sticky="w")
        width_entry = self._create_item_entry(row_frame, 170)
        width_entry.grid(row=0, column=1, padx=5, sticky="w")
        height_entry = self._create_item_entry(row_frame, 170)
        height_entry.grid(row=0, column=2, padx=5, sticky="w")
        count_entry = self._create_item_entry(row_frame, 130)
        count_entry.grid(row=0, column=3, padx=5, sticky="w")
        self._bind_enter(width_entry, height_entry)
        self._bind_enter(height_entry, count_entry)
        count_entry.bind(
            "<Return>",
            lambda event, index=item_number - 1: self._focus_after_item(
                event, index
            ),
            add="+",
        )
        self.blind_items.append(
            {
                "frame": row_frame,
                "width": width_entry,
                "height": height_entry,
                "count": count_entry,
            }
        )
        if len(self.blind_items) >= MAX_BLIND_ITEMS:
            self.add_blind_button.configure(state="disabled")
        if focus_new:
            width_entry.focus_set()

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

        items = []
        for item_number, fields in enumerate(self.blind_items, start=1):
            raw_values = (
                fields["width"].get().strip(),
                fields["height"].get().strip(),
                fields["count"].get().strip(),
            )
            if not any(raw_values):
                continue
            try:
                width_cm = parse_dimension(
                    raw_values[0], f"šířku žaluzie {item_number}"
                )
                height_cm = parse_dimension(
                    raw_values[1], f"výšku žaluzie {item_number}"
                )
                count = parse_count(raw_values[2])
            except ValueError as error:
                messagebox.showerror(
                    "Neplatné údaje", str(error), parent=self
                )
                return

            try:
                price = lookup_price(
                    self.blind_type.get(), width_cm * 10, height_cm * 10
                )
            except (FileNotFoundError, ValueError) as error:
                self.price_result.configure(text="")
                messagebox.showerror(
                    "Cenu se nepodařilo vypočítat",
                    f"Žaluzie {item_number}: {error}",
                    parent=self,
                )
                return
            items.append(
                {
                    "width_cm": width_cm,
                    "height_cm": height_cm,
                    "holder_count": calculate_holder_count(width_cm),
                    "count": count,
                    "price": price,
                }
            )

        if not items:
            messagebox.showerror(
                "Chybí žaluzie",
                "Vyplňte parametry alespoň jedné žaluzie.",
                parent=self,
            )
            return

        total_price = sum(item["price"] * item["count"] for item in items)
        self.price_result.configure(
            text=f"Celková cena: {total_price:,} Kč".replace(",", " ")
        )

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
                items,
                quote_number,
            )
        except Exception as error:
            messagebox.showerror(
                "Vytvoření ceníku se nezdařilo", str(error), parent=self
            )
            return

        messagebox.showinfo(
            "Ceník vytvořen", f"Nový ceník:\n{output_path}", parent=self
        )
        for entry in (self.entry_customer_name, self.entry_quote_number):
            entry.delete(0, "end")
        for item in self.blind_items[1:]:
            item["frame"].destroy()
        self.blind_items = self.blind_items[:1]
        for field in ("width", "height", "count"):
            self.blind_items[0][field].delete(0, "end")
        self.add_blind_button.configure(state="normal")
        self.price_result.configure(text="")
        self.blind_type.set(BLIND_TYPES[0])


def main() -> None:
    ctk.set_appearance_mode("light")
    ctk.set_default_color_theme("blue")
    ShadeInstallApp().mainloop()


if __name__ == "__main__":
    main()