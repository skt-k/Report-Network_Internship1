import tkinter as tk
from tkinter import ttk


def select_building(buildings):
    selected_building = []

    root = tk.Tk()
    root.title("เลือกตึก")
    root.geometry("320x140")
    root.resizable(False, False)
    root.eval("tk::PlaceWindow . center")

    tk.Label(
        root,
        text="เลือกตึกที่ต้องการ generate report:",
        font=("TH Sarabun New", 11),
    ).pack(pady=12)

    selected = tk.StringVar(value=buildings[0])
    dropdown = ttk.Combobox(
        root,
        textvariable=selected,
        values=buildings,
        state="readonly",
        width=28,
        font=("TH Sarabun New", 11),
    )
    dropdown.pack()

    def confirm():
        selected_building.append(selected.get())
        root.destroy()

    def on_close():
        root.destroy()

    root.protocol("WM_DELETE_WINDOW", on_close)
    tk.Button(
        root,
        text="ตกลง",
        command=confirm,
        width=12,
        font=("TH Sarabun New", 10),
    ).pack(pady=12)
    root.mainloop()

    return selected_building[0] if selected_building else None
