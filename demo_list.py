from pathlib import Path
import tkinter as tk
from tkinter import ttk


def find_demo_files(base_dir: Path) -> list[Path]:
    """Find Python demo files in the current folder, excluding this script."""
    current_script = Path(__file__).name
    return sorted(
        [path for path in base_dir.glob("*.py") if path.name != current_script],
        key=lambda p: p.name.lower(),
    )


def build_ui(demo_files: list[Path]) -> None:
    root = tk.Tk()
    root.title("Danh sach demo Python")
    root.geometry("520x340")
    root.minsize(420, 260)

    container = ttk.Frame(root, padding=12)
    container.pack(fill="both", expand=True)

    title = ttk.Label(container, text="Danh sach file demo Python")
    title.pack(anchor="w", pady=(0, 8))

    columns = ("stt", "ten_file")
    tree = ttk.Treeview(container, columns=columns, show="headings")
    tree.heading("stt", text="STT")
    tree.heading("ten_file", text="Ten file")
    tree.column("stt", width=70, anchor="center", stretch=False)
    tree.column("ten_file", width=380, anchor="w", stretch=True)

    scrollbar = ttk.Scrollbar(container, orient="vertical", command=tree.yview)
    tree.configure(yscrollcommand=scrollbar.set)

    tree.pack(side="left", fill="both", expand=True)
    scrollbar.pack(side="right", fill="y")

    if demo_files:
        for index, file_path in enumerate(demo_files, start=1):
            tree.insert("", "end", values=(index, file_path.name))
    else:
        tree.insert("", "end", values=("-", "Khong tim thay file demo nao."))

    root.mainloop()


def main() -> None:
    base_dir = Path(__file__).resolve().parent
    demo_files = find_demo_files(base_dir)
    build_ui(demo_files)


if __name__ == "__main__":
    main()
