"""Minimální bezokenní GUI základ aplikace ShadeInstall."""

import customtkinter as ctk


class ShadeInstallApp:
    """Hlavní okno a prostor pro budoucí formulář aplikace."""

    def __init__(self, root: ctk.CTk) -> None:
        self.root = root
        self.root.title("ShadeInstall - Interní aplikace")
        self.root.geometry("800x600")
        self.root.configure(fg_color="white")

        # Sem později patří formulář a ovládací prvky; logiku Excelu
        # oddělte tak, aby zapisovala jen do určených buněk ceníku.
        self.main_frame = ctk.CTkFrame(root, fg_color="white", corner_radius=0)
        self.main_frame.pack(fill="both", expand=True)


def main() -> None:
    ctk.set_appearance_mode("light")
    root = ctk.CTk()
    ShadeInstallApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()