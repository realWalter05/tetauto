import tkinter as tk
from tkinter import ttk
from tkinter import filedialog as fd
from tkinter.messagebox import showinfo
from excel import Excel


def select_file():
    filetypes = (
        ('Excel soubory', '*.xlsx'),
        ('All files', '*.*')
    )

    filename = fd.askopenfilename(
        title='Vyberte excel soubor',
        initialdir='/',
        filetypes=filetypes)

    Excel(filename)

    showinfo(
        title='Status',
        message="Soubor s dohody byl vytvořen na ploše. Můžete zavřít okno."
    )


def setup():
    # create the root window
    root = tk.Tk()
    root.title('Vyberte excel soubor')
    root.resizable(False, False)
    root.geometry('300x150')

    # open button
    open_button = ttk.Button(
        root,
        text='Vyberte excel soubor',
        command=select_file
    )

    label = tk.Label(text="Auto-tvorba dohod", font=("Helvetica", 18))
    label.pack(expand=True)

    open_button.pack(expand=True)
    return root




if __name__ == "__main__":
    # Run the app :)
    setup().mainloop()
