import tkinter as tk

from app import CertificateApp


if __name__ == "__main__":
    root = tk.Tk()
    root.attributes('-alpha', 0)
    root.configure(bg='#f0f0f0')
    app = CertificateApp(root)
    root.mainloop()
