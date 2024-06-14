import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
import pyodbc
import pandas as pd


def select_file():
    file_path = filedialog.askopenfilename(
        filetypes=[("Access Database Files", "*.accdb *.mdb")])
    entry_file_path.delete(0, tk.END)
    entry_file_path.insert(0, file_path)


def generate_csv():
    file_path = entry_file_path.get()
    query = text_query.get("1.0", tk.END).strip()

    if not file_path or not query:
        messagebox.showerror("Error", "Please provide both file path and SQL query.")
        return

    try:
        conn_str = f'DRIVER={{Microsoft Access Driver (*.mdb, *.accdb)}};DBQ={file_path};'
        conn = pyodbc.connect(conn_str)
        data = pd.read_sql(query, conn)
        conn.close()

        save_path = filedialog.asksaveasfilename(defaultextension=".csv",
                                                 filetypes=[("CSV Files", "*.csv")])
        if save_path:
            data.to_csv(save_path, index=False)
            messagebox.showinfo("Success", f"CSV file has been saved to {save_path}")

    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {str(e)}")


# Create the main window
root = tk.Tk()
root.title("Access Database to CSV")
root.geometry("600x400")

# Create and place the widgets
frame = ttk.Frame(root, padding="10")
frame.pack(fill=tk.BOTH, expand=True)

label_file_path = ttk.Label(frame, text="Access Database File:")
label_file_path.grid(row=0, column=0, sticky=tk.W, pady=5)

entry_file_path = ttk.Entry(frame, width=50)
entry_file_path.grid(row=0, column=1, pady=5, sticky=tk.EW)

button_browse = ttk.Button(frame, text="Browse...", command=select_file)
button_browse.grid(row=0, column=2, padx=5, pady=5)

label_query = ttk.Label(frame, text="SQL Query:")
label_query.grid(row=1, column=0, sticky=tk.NW, pady=5)

text_query = tk.Text(frame, width=50, height=10)
text_query.grid(row=1, column=1, columnspan=2, pady=5, sticky=tk.EW)

button_generate = ttk.Button(frame, text="Generate CSV", command=generate_csv)
button_generate.grid(row=2, column=1, pady=10)

# Configure column weights for better resizing behavior
frame.columnconfigure(1, weight=1)

# Add some padding around the widgets
for widget in frame.winfo_children():
    widget.grid_configure(padx=5, pady=5)

# Add a style for better appearance
style = ttk.Style()
style.configure("TButton", padding=6, relief="flat", background="#ccc")
style.configure("TLabel", padding=6)

# Run the application
root.mainloop()
