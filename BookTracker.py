import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import pyodbc
import datetime
import os

# --- Database Connection ---
CONN_STR = (
    'DRIVER={ODBC Driver 17 for SQL Server};'
    'SERVER=SERVERHOSTNAME;DATABASE=BookDB;UID=USERNAME;PWD=PASSWORD'
)

def get_conn():
    return pyodbc.connect(CONN_STR)

def fetch_books(search_term=""):
    conn = get_conn()
    cursor = conn.cursor()
    if search_term:
        cursor.execute("SELECT id, title, author, date_added, summary, rating FROM Books WHERE title LIKE ?", ('%' + search_term + '%',))
    else:
        cursor.execute("SELECT id, title, author, date_added, summary, rating FROM Books")
    rows = cursor.fetchall()
    conn.close()
    return rows

def add_book(title, author, summary, rating):
    conn = get_conn()
    cursor = conn.cursor()
    cursor.execute(
        "INSERT INTO Books (title, author, date_added, summary, rating) VALUES (?, ?, ?, ?, ?)",
        (title, author, datetime.datetime.now(), summary, rating)
    )
    conn.commit()
    conn.close()

def update_book(book_id, title, author, summary, rating):
    conn = get_conn()
    cursor = conn.cursor()
    cursor.execute(
        "UPDATE Books SET title=?, author=?, summary=?, rating=? WHERE id=?",
        (title, author, summary, rating, book_id)
    )
    conn.commit()
    conn.close()

def delete_book(book_id):
    conn = get_conn()
    cursor = conn.cursor()
    cursor.execute("DELETE FROM Books WHERE id=?", (book_id,))
    conn.commit()
    conn.close()

class BookTrackerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Book Tracker")
        self.root.configure(bg="#f4f4f4")

        # --- Styles ---
        style = ttk.Style()
        style.theme_use('clam')
        style.configure("Treeview", background="#f9f9fa", fieldbackground="#f9f9fa", foreground="#333")
        style.configure("Treeview.Heading", background="#e0e0e0", foreground="#333", font=('Arial', 10, 'bold'))

        # --- Variables ---
        self.title_var = tk.StringVar()
        self.author_var = tk.StringVar()
        self.summary_var = tk.StringVar()
        self.rating_var = tk.StringVar()
        self.search_var = tk.StringVar()

        # --- Search Bar ---
        search_frame = tk.Frame(root, bg="#f4f4f4")
        search_frame.pack(fill="x", padx=12, pady=(10, 0))
        tk.Label(search_frame, text="Search Title:", bg="#f4f4f4", font=("Arial", 10)).pack(side="left", padx=(0,5))
        search_entry = tk.Entry(search_frame, textvariable=self.search_var, width=30, font=("Arial", 10))
        search_entry.pack(side="left", padx=(0,5))
        tk.Button(search_frame, text="Search", bg="#a084e8", fg="white", font=("Arial", 10, "bold"),
                  command=self.search).pack(side="left", padx=2)
        tk.Button(search_frame, text="Show All", bg="#bdbdbd", fg="black", font=("Arial", 10, "bold"),
                  command=self.refresh).pack(side="left", padx=2)

        # --- Data Entry Frame ---
        entry_frame = tk.LabelFrame(root, text="Book Details", bg="#f4f4f4", font=("Arial", 10, "bold"), bd=2)
        entry_frame.pack(fill="x", padx=12, pady=10)

        tk.Label(entry_frame, text="Title:", bg="#f4f4f4", font=("Arial", 10)).grid(row=0, column=0, sticky="e", pady=4, padx=(0,5))
        tk.Entry(entry_frame, textvariable=self.title_var, width=30, font=("Arial", 10)).grid(row=0, column=1, pady=4, sticky="w")

        tk.Label(entry_frame, text="Author:", bg="#f4f4f4", font=("Arial", 10)).grid(row=1, column=0, sticky="e", pady=4, padx=(0,5))
        tk.Entry(entry_frame, textvariable=self.author_var, width=30, font=("Arial", 10)).grid(row=1, column=1, pady=4, sticky="w")

        tk.Label(entry_frame, text="Summary:", bg="#f4f4f4", font=("Arial", 10)).grid(row=2, column=0, sticky="e", pady=4, padx=(0,5))
        tk.Entry(entry_frame, textvariable=self.summary_var, width=30, font=("Arial", 10)).grid(row=2, column=1, pady=4, sticky="w")

        tk.Label(entry_frame, text="Rating:", bg="#f4f4f4", font=("Arial", 10)).grid(row=3, column=0, sticky="e", pady=4, padx=(0,5))
        tk.Entry(entry_frame, textvariable=self.rating_var, width=10, font=("Arial", 10)).grid(row=3, column=1, pady=4, sticky="w")

        # --- Button Frame ---
        button_frame = tk.Frame(entry_frame, bg="#f4f4f4")
        button_frame.grid(row=0, column=2, rowspan=4, padx=(20,0), pady=2, sticky="ns")

        tk.Button(button_frame, text="Add", bg="#4caf50", fg="white", font=("Arial", 10, "bold"),
                  width=10, command=self.add).pack(pady=2)
        tk.Button(button_frame, text="Update", bg="#2196f3", fg="white", font=("Arial", 10, "bold"),
                  width=10, command=self.update).pack(pady=2)
        tk.Button(button_frame, text="Delete", bg="#f44336", fg="white", font=("Arial", 10, "bold"),
                  width=10, command=self.delete).pack(pady=2)
        tk.Button(button_frame, text="Print Results", bg="#ff9800", fg="white", font=("Arial", 10, "bold"),
                  width=12, command=self.print_results).pack(pady=2)
        # --- Refresh Button ---
        tk.Button(button_frame, text="Refresh", bg="#8bc34a", fg="black", font=("Arial", 10, "bold"),
                  width=10, command=self.refresh).pack(pady=2)

        # --- Table Frame ---
        table_frame = tk.Frame(root, bg="#f4f4f4")
        table_frame.pack(fill="both", expand=True, padx=12, pady=(0,12))

        columns = ("ID", "Title", "Author", "Date", "Summary", "Rating")
        self.tree = ttk.Treeview(table_frame, columns=columns, show="headings", height=12)
        for col in columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, anchor="center", width=100 if col=="ID" else 140)
        self.tree.column("Summary", width=200)
        self.tree.pack(side="left", fill="both", expand=True)

        # Add vertical scrollbar
        vsb = ttk.Scrollbar(table_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        vsb.pack(side="right", fill="y")

        self.tree.bind('<<TreeviewSelect>>', self.on_select)

        self.refresh()

    def refresh(self):
        for row in self.tree.get_children():
            self.tree.delete(row)
        for row in fetch_books():
            row = list(row)
            # Format date_added (index 3)
            if isinstance(row[3], datetime.datetime):
                row[3] = row[3].strftime("%Y-%m-%d %H:%M")
            self.tree.insert("", "end", values=row)

    def add(self):
        if not self.title_var.get() or not self.author_var.get():
            messagebox.showwarning("Missing Data", "Title and Author are required.")
            return
        add_book(self.title_var.get(), self.author_var.get(), self.summary_var.get(), self.rating_var.get())
        self.clear_fields()
        self.refresh()

    def update(self):
        selected = self.tree.selection()
        if not selected:
            messagebox.showwarning("Select a book", "No book selected.")
            return
        book_id = self.tree.item(selected[0])['values'][0]
        update_book(book_id, self.title_var.get(), self.author_var.get(), self.summary_var.get(), self.rating_var.get())
        self.clear_fields()
        self.refresh()

    def delete(self):
        selected = self.tree.selection()
        if not selected:
            messagebox.showwarning("Select a book", "No book selected.")
            return
        book_id = self.tree.item(selected[0])['values'][0]
        if messagebox.askyesno("Confirm Delete", "Are you sure you want to delete this book?"):
            delete_book(book_id)
            self.clear_fields()
            self.refresh()

    def search(self):
        term = self.search_var.get()
        for row in self.tree.get_children():
            self.tree.delete(row)
        for row in fetch_books(term):
            row = list(row)
            if isinstance(row[3], datetime.datetime):
                row[3] = row[3].strftime("%Y-%m-%d %H:%M")
            self.tree.insert("", "end", values=row)

    def print_results(self):
        rows = [self.tree.item(child)['values'] for child in self.tree.get_children()]
        if not rows:
            messagebox.showinfo("No Data", "There are no results to print.")
            return
        # Ask user where to save the file
        file_path = filedialog.asksaveasfilename(
            defaultextension=".txt",
            filetypes=[("Text Files", "*.txt"), ("All Files", "*.*")],
            title="Save Book Search Results"
        )
        if not file_path:
            return  # User cancelled
        try:
            with open(file_path, "w", encoding="utf-8") as f:
                for row in rows:
                    f.write(" | ".join(str(item) for item in row) + "\n")
            messagebox.showinfo("Printed", f"Results printed to {file_path}.\nPlease print the file manually.")
        except Exception as e:
            messagebox.showerror("Error", f"Could not save file:\n{e}")

    def on_select(self, event):
        selected = self.tree.selection()
        if not selected:
            return
        values = self.tree.item(selected[0])['values']
        self.title_var.set(values[1])
        self.author_var.set(values[2])
        self.summary_var.set(values[4])
        self.rating_var.set(values[5])

    def clear_fields(self):
        self.title_var.set("")
        self.author_var.set("")
        self.summary_var.set("")
        self.rating_var.set("")

if __name__ == "__main__":
    root = tk.Tk()
    app = BookTrackerApp(root)
    root.mainloop()