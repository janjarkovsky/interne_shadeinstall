"""ShadeInstall - kalkulátor plochy žaluzií a Excel zápis."""

import sys
from datetime import datetime
from decimal import Decimal, InvalidOperation
from pathlib import Path
from tkinter import filedialog, messagebox

import customtkinter as ctk
from openpyxl import load_workbook
from openpyxl.styles import Font
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


def write_calculation_to_workbook(
    source_path: Path,
    width_mm: Decimal,
    height_mm: Decimal,
) -> Path:
    """Uloží nový záznam do nové kopie sešitu, původní listy nemění."""
    workbook = load_workbook(
        source_path,
        data_only=False,
        keep_vba=source_path.suffix.lower() == ".xlsm",
    )
    sheet_name = "Výpočty ShadeInstall"
    headers = ("Datum", "Šířka (mm)", "Výška (mm)", "Plocha (m²)")
    try:
        if sheet_name in workbook.sheetnames:
            worksheet = workbook[sheet_name]
            existing_headers = tuple(
                worksheet.cell(row=1, column=col).value for col in range(1, 5)
            )
            if all(value is None for value in existing_headers):
                worksheet.append(headers)
            elif existing_headers != headers:
                raise ValueError(f"List '{sheet_name}' má jinou strukturu.")
        else:
            worksheet = workbook.create_sheet(sheet_name)
            worksheet.append(headers)

        for cell in worksheet[1]:
            cell.font = Font(bold=True)
        worksheet.append(
            [
                datetime.now(),
                float(width_mm),
                float(height_mm),
                float(width_mm * height_mm / Decimal("1000000")),
            ]
        )
        row = worksheet.max_row
        worksheet.cell(row=row, column=1).number_format = "yyyy-mm-dd hh:mm:ss"
        worksheet.cell(row=row, column=4).number_format = "0.0000"
        worksheet.freeze_panes = "A2"

        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
        output_path = source_path.with_name(
            f"{source_path.stem}_ShadeInstall_{timestamp}{source_path.suffix}"
        )
        workbook.save(output_path)
        return output_path
    finally:
        workbook.close()


class ShadeInstallApp(ctk.CTk):
    def __init__(self) -> None:
        super().__init__()
        self.title("ShadeInstall.cz - Kalkulátor žaluzií")
        self.geometry("1050x700")
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
        ctk.CTkLabel(
            content,
            text="Parametry žaluzie",
            font=("Segoe UI", 16, "bold"),
            text_color="#1e293b",
        ).pack(anchor="w", padx=30, pady=(30, 15))

        self.entry_width = self._add_dimension_field(content, "Šířka (mm):")
        self.entry_height = self._add_dimension_field(content, "Výška (mm):")
        ctk.CTkButton(
            content,
            text="Spočítat a zapsat do ceníku",
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
    def _add_dimension_field(parent: ctk.CTkFrame, label: str) -> ctk.CTkEntry:
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
        try:
            width_mm = parse_dimension(self.entry_width.get(), "šířku")
            height_mm = parse_dimension(self.entry_height.get(), "výšku")
        except ValueError as error:
            messagebox.showerror("Neplatné rozměry", str(error), parent=self)
            return

        selected_file = filedialog.askopenfilename(
            title="Vyberte Excel ceník",
            filetypes=(("Excel sešit", "*.xlsx *.xlsm"),),
            parent=self,
        )
        if not selected_file:
            return
        source_path = Path(selected_file)
        try:
            output_path = write_calculation_to_workbook(
                source_path, width_mm, height_mm
            )
        except Exception as error:
            messagebox.showerror(
                "Zápis do Excelu se nezdařil",
                f"Původní sešit nebyl změněn.\n\n{error}",
                parent=self,
            )
            return

        area_m2 = width_mm * height_mm / Decimal("1000000")
        messagebox.showinfo(
            "Záznam uložen",
            f"Plocha: {area_m2:.4f} m²\n\nKopie ceníku:\n{output_path}",
            parent=self,
        )


def main() -> None:
    ctk.set_appearance_mode("light")
    ctk.set_default_color_theme("blue")
    ShadeInstallApp().mainloop()


if __name__ == "__main__":
    main()