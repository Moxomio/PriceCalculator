import tkinter as tk
from tkinter import ttk, messagebox
import pandas as pd

# Load Excel data
EXCEL_FILE = "product_database.xlsx"
xls = pd.ExcelFile(EXCEL_FILE)

# Load profile sheets
sheet_names = [s for s in xls.sheet_names if s.upper().startswith("PROFILE")]
sheets_dict = {name: xls.parse(name) for name in sheet_names}
for name, df in sheets_dict.items():
    df["Width(mm)"] = pd.to_numeric(df["Width(mm)"], errors="coerce")
    df["Length(mm)"] = pd.to_numeric(df["Length(mm)"], errors="coerce")
    df["Type"] = df["Type"].astype(str).str.strip().str.upper()
    df["Price(RM)"] = df["Price(RM)"].astype(str).str.replace("/M", "").str.strip()
    df["Total"] = df["Total"].astype(str).str.strip()
    sheets_dict[name] = df

# Load accessory sheet
accessory_df = xls.parse("ACCESSORY")
accessory_df["Quantity"] = accessory_df["Quantity"].astype(str).str.strip()
accessory_df["Price"] = pd.to_numeric(accessory_df["Price"], errors="coerce")
accessory_df["Type"] = accessory_df["Type"].astype(str).str.strip()
accessory_df["DESCRIPTION"] = accessory_df["DESCRIPTION"].astype(str).str.strip()

# App base
root = tk.Tk()
root.title("Product Calculator")
root.geometry("800x960")

# Navigation
def show_frame(frame):
    frame.tkraise()

container = tk.Frame(root)
container.pack(side="top", fill="both", expand=True)
frames = {}

for F in ["Main", "Profile", "Accessory"]:
    frame = tk.Frame(container)
    frame.grid(row=0, column=0, sticky="nsew")
    frames[F] = frame

# === Main Menu ===
tk.Label(frames["Main"], text="Select Item Type", font=("Arial", 18)).pack(pady=30)
tk.Button(frames["Main"], text="Profile", width=20, height=2, command=lambda: show_frame(frames["Profile"])).pack(pady=10)
tk.Button(frames["Main"], text="Accessory", width=20, height=2, command=lambda: show_frame(frames["Accessory"])).pack(pady=10)

# === Profile Page ===
tk.Label(frames["Profile"], text="PROFILE CALCULATION", font=("Arial", 16)).grid(row=0, columnspan=2, pady=10)

sheet_var = tk.StringVar()
width_var = tk.StringVar()
length_var = tk.StringVar()
type_var = tk.StringVar()

tk.Label(frames["Profile"], text="Sheet").grid(row=1, column=0, sticky='w')
sheet_dropdown = ttk.Combobox(frames["Profile"], textvariable=sheet_var, values=sheet_names, state="readonly")
sheet_dropdown.grid(row=1, column=1, padx=10, pady=5)

tk.Label(frames["Profile"], text="Width (mm)").grid(row=2, column=0, sticky='w')
width_dropdown = ttk.Combobox(frames["Profile"], textvariable=width_var, state="readonly")
width_dropdown.grid(row=2, column=1, padx=10, pady=5)

tk.Label(frames["Profile"], text="Length (mm)").grid(row=3, column=0, sticky='w')
length_dropdown = ttk.Combobox(frames["Profile"], textvariable=length_var, state="readonly")
length_dropdown.grid(row=3, column=1, padx=10, pady=5)

tk.Label(frames["Profile"], text="Type").grid(row=4, column=0, sticky='w')
type_dropdown = ttk.Combobox(frames["Profile"], textvariable=type_var, state="readonly")
type_dropdown.grid(row=4, column=1, padx=10, pady=5)

p_total = tk.Entry(frames["Profile"])
p_qty = tk.Entry(frames["Profile"])
p_holes = tk.Entry(frames["Profile"])

tk.Label(frames["Profile"], text="Total (mm)").grid(row=5, column=0, sticky='w')
p_total.grid(row=5, column=1)

tk.Label(frames["Profile"], text="Quantity (pcs)").grid(row=6, column=0, sticky='w')
p_qty.grid(row=6, column=1)

tk.Label(frames["Profile"], text="Holes (0 if none)").grid(row=7, column=0, sticky='w')
p_holes.grid(row=7, column=1)

profile_result = tk.Text(frames["Profile"], height=15, width=80)
profile_result.grid(row=9, columnspan=2, pady=10)

def update_profile_options(*args):
    df = sheets_dict.get(sheet_var.get())
    if df is not None:
        width_dropdown["values"] = sorted(df["Width(mm)"].dropna().unique())
        length_dropdown["values"] = sorted(df["Length(mm)"].dropna().unique())
        type_dropdown["values"] = sorted(df["Type"].dropna().unique())
        if width_dropdown["values"]: width_var.set(width_dropdown["values"][0])
        if length_dropdown["values"]: length_var.set(length_dropdown["values"][0])
        if type_dropdown["values"]: type_var.set(type_dropdown["values"][0])

sheet_var.trace_add("write", update_profile_options)

def calculate_profile():
    profile_result.delete("1.0", tk.END)
    df = sheets_dict.get(sheet_var.get())
    try:
        w = int(width_var.get())
        l = float(length_var.get())
        t = type_var.get().strip().upper()
        total = float(p_total.get())
        qty = int(p_qty.get())
        holes = int(p_holes.get())
    except:
        messagebox.showerror("Input Error", "Please fill all profile fields correctly.")
        return
    filtered = df[(df["Width(mm)"] == w) & (df["Length(mm)"] == l) & (df["Type"] == t)]
    if filtered.empty:
        profile_result.insert(tk.END, "No matching profile found.\n")
        return
    total_m = (total * qty) / 1000
    matched_price = None
    for _, row in filtered.iterrows():
        try:
            val = float(row["Total"].strip("<>=M "))
            price = float(row["Price(RM)"])
            bracket = row["Total"].strip()
            if bracket.startswith("<") and total_m < val:
                matched_price = price
                break
            elif bracket.startswith(">") and total_m > val:
                matched_price = price
            elif bracket.startswith("=") and total_m == val:
                matched_price = price
                break
        except: continue
    if matched_price is None:
        profile_result.insert(tk.END, "No matching price found.\n")
        return
    m_cost = matched_price * total_m
    c_cost = qty * 2 * 3
    h_cost = holes * 8
    total_cost = m_cost + c_cost + h_cost
    profile_result.insert(tk.END, f"Material Cost: RM {m_cost:.2f}\nCutting: RM {c_cost:.2f}\nHole: RM {h_cost:.2f}\n")
    profile_result.insert(tk.END, f"\nFinal Total: RM {total_cost:.2f}\n", "big")
    profile_result.tag_config("big", font=("Arial", 14, "bold"))


tk.Button(frames["Profile"], text="Calculate", command=calculate_profile).grid(row=8, columnspan=2, pady=5)
tk.Button(frames["Profile"], text="← Back", command=lambda: show_frame(frames["Main"])).grid(row=10, column=0, pady=5)

# === Accessory Page ===
tk.Label(frames["Accessory"], text="ACCESSORY CALCULATION", font=("Arial", 16)).grid(row=0, columnspan=2, pady=10)

a_type = tk.StringVar()
a_desc = tk.StringVar()
a_qty = tk.Entry(frames["Accessory"])

tk.Label(frames["Accessory"], text="Accessory Type").grid(row=1, column=0, sticky='w')
type_box = ttk.Combobox(frames["Accessory"], textvariable=a_type, values=sorted(accessory_df["Type"].unique()), state="readonly")
type_box.grid(row=1, column=1)

tk.Label(frames["Accessory"], text="Description").grid(row=2, column=0, sticky='w')
desc_box = ttk.Combobox(frames["Accessory"], textvariable=a_desc, state="readonly")
desc_box.grid(row=2, column=1)

tk.Label(frames["Accessory"], text="Quantity").grid(row=3, column=0, sticky='w')
a_qty.grid(row=3, column=1)

accessory_result = tk.Text(frames["Accessory"], height=10, width=80)
accessory_result.grid(row=5, columnspan=2, pady=10)

def update_desc(*args):
    df = accessory_df[accessory_df["Type"] == a_type.get()]
    desc_box["values"] = sorted(df["DESCRIPTION"].unique())
    if not df.empty:
        a_desc.set(df["DESCRIPTION"].iloc[0])

a_type.trace_add("write", update_desc)

def calculate_accessory():
    accessory_result.delete("1.0", tk.END)
    try:
        qty = int(a_qty.get())
    except:
        messagebox.showerror("Input Error", "Enter a valid quantity.")
        return
    df = accessory_df[(accessory_df["Type"] == a_type.get()) & (accessory_df["DESCRIPTION"] == a_desc.get())]
    price = None
    for _, row in df.iterrows():
        bracket = row["Quantity"].strip()
        try:
            val = int(bracket.strip("<>= "))
            if bracket.startswith("<") and qty < val:
                price = row["Price"]
                break
            elif bracket.startswith(">") and qty > val:
                price = row["Price"]
            elif bracket.startswith("=") and qty == val:
                price = row["Price"]
                break
        except: continue
    if price is None:
        accessory_result.insert(tk.END, "No matching price found.\n")
        return
    total = price * qty
    accessory_result.insert(tk.END, f"Accessory: {a_desc.get()}\nQuantity: {qty}\nUnit Price: RM {price:.2f}\n")
    accessory_result.insert(tk.END, f"\nFinal Total: RM {total:.2f}\n", "big")
    accessory_result.tag_config("big", font=("Arial", 14, "bold"))


tk.Button(frames["Accessory"], text="Calculate", command=calculate_accessory).grid(row=4, columnspan=2, pady=5)
tk.Button(frames["Accessory"], text="← Back", command=lambda: show_frame(frames["Main"])).grid(row=6, column=0, pady=5)

# Start UI
show_frame(frames["Main"])
root.mainloop()
