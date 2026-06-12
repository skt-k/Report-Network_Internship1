import tkinter as tk
from tkinter import ttk
import threading


def select_building(buildings, on_generate_callback=None):
    selected_building = [buildings[0]]  # ตั้งค่าเริ่มต้น
    generating = [False]

    root = tk.Tk()
    root.title("Network Report Generator")
    root.geometry("420x380")
    root.resizable(False, False)
    root.eval("tk::PlaceWindow . center")

    tk.Label(
        root,
        text="เลือกตึกที่ต้องการ generate report:",
        font=("TH Sarabun New", 11),
        justify="left",
    ).pack(pady=10, padx=10, anchor="w")

    selected = tk.StringVar(value=buildings[0])
    dropdown = ttk.Combobox(
        root,
        textvariable=selected,
        values=buildings,
        state="readonly",
        width=40,
        font=("TH Sarabun New", 11),
    )
    dropdown.pack(padx=10, pady=5)

    # สถานะการทำงาน
    tk.Label(
        root,
        text="สถานะการทำงาน:",
        font=("TH Sarabun New", 10, "bold"),
    ).pack(pady=(15, 5), padx=10, anchor="w")

    status_frame = tk.Frame(root, relief="solid", borderwidth=1, bg="white")
    status_frame.pack(padx=10, pady=5, fill="both", expand=True)

    status_text = tk.Text(
        status_frame,
        height=10,
        width=45,
        font=("Courier New", 9),
        bg="white",
        fg="black",
        state="disabled",
    )
    status_text.pack(padx=5, pady=5, fill="both", expand=True)

    def update_status(message):
        status_text.config(state="normal")
        status_text.insert("end", message + "\n")
        status_text.see("end")
        status_text.config(state="disabled")
        root.update()

    def clear_status():
        status_text.config(state="normal")
        status_text.delete("1.0", "end")
        status_text.config(state="disabled")

    def confirm():
        if generating[0]:
            return
        
        selected_building[0] = selected.get()
        clear_status()
        update_status(f"🔄 กำลังสร้างรายงานสำหรับ: {selected_building[0]}...")
        generating[0] = True
        dropdown.config(state="disabled")
        confirm_btn.config(state="disabled")

        def run_generation():
            try:
                if on_generate_callback:
                    on_generate_callback(selected_building[0], update_status)
            except Exception as e:
                update_status(f"❌ เกิดข้อผิดพลาด: {str(e)}")
            finally:
                generating[0] = False
                dropdown.config(state="readonly")
                confirm_btn.config(state="normal")

        # รัน generate ใน thread เพื่อไม่ให้ GUI ค้าง
        thread = threading.Thread(target=run_generation, daemon=True)
        thread.start()

    def on_close():
        if not generating[0]:
            root.destroy()

    button_frame = tk.Frame(root)
    button_frame.pack(pady=10)

    confirm_btn = tk.Button(
        button_frame,
        text="Generate",
        command=confirm,
        width=12,
        font=("TH Sarabun New", 10),
        bg="#006E04",
        fg="white",
    )
    confirm_btn.pack(side="left", padx=5)

    tk.Button(
        button_frame,
        text="Cancel",
        command=on_close,
        width=12,
        font=("TH Sarabun New", 10),
    ).pack(side="left", padx=5)

    root.protocol("WM_DELETE_WINDOW", on_close)
    root.mainloop()

    return selected_building[0]
