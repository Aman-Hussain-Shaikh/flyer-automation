import sys
import tkinter
import customtkinter

def verify_installation():
    print("Python Version:", sys.version)
    print("Tkinter Version: Available" if tkinter else "Tkinter: Not Available")
    print("CustomTkinter Version:", customtkinter.__version__)

    # Basic GUI Test
    root = customtkinter.CTk()
    root.geometry("400x300")
    root.title("CustomTkinter Verification")

    label = customtkinter.CTkLabel(root, text="Installation Successful!")
    label.pack(pady=20)

    root.after(3000, root.destroy)  # Auto-close after 3 seconds
    root.mainloop()

if __name__ == "__main__":
    verify_installation()