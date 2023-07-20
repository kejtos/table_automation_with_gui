import os
from functools import partial
from dataclasses import dataclass

from tkinter import Tk, Canvas, Text, Button, Label, Frame, Scrollbar
from tkinter import filedialog, messagebox
from tkinter import SUNKEN, RAISED

from main import import_tables, create_table


parent_path = os.getcwd()
inputs_path = fr'{parent_path}\vstupy'
try:
    input_folders = os.listdir(inputs_path)
except:
    messagebox.showerror('Error!', f'generator.exe musí být ve stejné složce, jako složka vstupy.')
initialdir = fr'{parent_path}\vstupy\{max(input_folders)}'
ICONPATH = fr'{parent_path}\icon.ico'
path_names = ('komponenty', 'matice_vyroby', 'vyroba', 'zavody', 'výstup')
basenames = ('komponenty.csv', 'matice_vyroby.txt', 'vyroba.txt', 'zavody.csv')
outputdir = fr'{parent_path}\prehledy'
label_texts = ('s komponentami', 's maticí výroby', 's výrobou', 'se závody')


@dataclass
class GlobalData:
    initialdir: str
    path_names: tuple[str]
    outputdir: str
    basenames: tuple[str]
    file_paths = None

    def __post_init__(self):
        if self.file_paths is None:
            self.file_paths = {key: fr'{initialdir}\{name}' for key, name in zip(self.path_names[0:-1], self.basenames)}


def get_curr_screen_geometry():
    root = Tk()
    root.update_idletasks()
    root.attributes('-fullscreen', True)
    root.state('iconic')
    geometry = root.winfo_geometry()
    root.destroy()
    return geometry


def browse_files(text_box, index, file):
    if file:
        file_path = filedialog.askopenfilename(initialdir=d.initialdir, title='Vyber soubor', filetypes=(('all files', '*.*'), ('csv files', '*.csv*'), ('text files', '*.txt*')))
        filename = os.path.basename(file_path)
        text_box.configure(state='normal')
        text_box.delete('1.0', 'end')
        text_box.insert('1.0', filename)
        text_box.configure(state='disabled')
        if file_path:
            d.initialdir = os.path.dirname(file_path)
            d.file_paths[d.basenames[index]] = file_path
    else:
        folder_path = filedialog.askdirectory(initialdir=d.initialdir, title='Vyber složku')
        d.initialdir = folder_path
        d.outputdir = folder_path
        show_path = f'.\{os.sep.join(os.path.normpath(folder_path).split(os.sep)[-3:])}'
        text_box.configure(state='normal')
        text_box.delete('1.0', 'end')
        text_box.insert('1.0', show_path)
        text_box.configure(state='disabled')


def run_create_table():
    matice_vyroby, komponenty, vyroba, zavody = import_tables(komponenty_path=d.file_paths['komponenty'], matice_vyroby_path=d.file_paths['matice_vyroby'], vyroba_path=d.file_paths['vyroba'], zavody_path=d.file_paths['zavody'])
    wrong_cols_comp = [col for col in ('ID_komponenty', 'Porizovaci_cena') if col not in komponenty.columns]
    wrong_cols_mat = [col for col in ('ID_produktu', 'ID_komponenty', 'Mnozstvi') if col not in matice_vyroby.columns]
    wrong_cols_vyr = [col for col in ('ID_produktu', 'ID_zavodu', 'Datum', 'Mnozstvi') if col not in vyroba.columns]
    wrong_cols_zav = [col for col in ('ID_zavodu', 'Misto') if col not in zavody.columns]
    wrong_cols = ((wrong_cols_comp, d.file_paths['komponenty']), (wrong_cols_mat, d.file_paths['matice_vyroby']), (wrong_cols_vyr, d.file_paths['vyroba']), (wrong_cols_zav, d.file_paths['zavody']))

    for cols, path in wrong_cols:
        if len(cols) > 0:
            if len(cols) > 1:
                messagebox.showerror('Error!', f'V souboru {os.path.basename(path)} chybí sloupec {cols[0]}, nebo je špatně pojmenován.')
            else:
                messagebox.showerror('Error!', f'V souboru {os.path.basename(path)} chybí sloupce {", ".join(cols)}, nebo jsou špatně pojmenované.')

    create_table(matice_vyroby=matice_vyroby, vyroba=vyroba, komponenty=komponenty, zavody=zavody, vystup=d.outputdir)
    labelos.configure(text=f'Přehled vytvořen ve složce {os.path.basename(d.outputdir)}!')


def on_mousewheel(event):
    if main_canvas.winfo_height() < main_frame.winfo_height():
        main_canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")


def press_create(button):
    button.config(relief=SUNKEN)
    r.update()
    r.after(100)
    button.config(relief=RAISED)
    run_create_table()

d = GlobalData(initialdir=initialdir, path_names=path_names, outputdir=outputdir, basenames=basenames)

r = Tk()#(theme='scidgrey')
r.lift()
r.title('Generátor přehledu výroby')
geometry = get_curr_screen_geometry()
cut = geometry.partition("x")
screen_width = int(cut[0])
screen_height = int(cut[2].partition("+")[0])
WIDTH = 790
HEIGHT = 260
X = round((screen_width - WIDTH) / 2)
Y = round((screen_height - HEIGHT) / 2)

r.geometry(f'{WIDTH}x{HEIGHT}+{X}+{Y}')

try:
    r.iconbitmap(ICONPATH)
except:
    pass

r.config(background='white')

main_frame = Frame(r,)
main_frame.pack(fill='both', expand=1)

main_canvas = Canvas(main_frame, highlightthickness=0)

h_sb = Scrollbar(main_frame, orient='horizontal', command=main_canvas.xview)
v_sb = Scrollbar(main_frame, orient='vertical', command=main_canvas.yview)
h_sb.pack(side='bottom', fill='x')
v_sb.pack(side='right', fill='y')

main_canvas.pack(side='left', fill='both', expand=1)

main_canvas.configure(yscrollcommand=v_sb.set, xscrollcommand=h_sb.set)
main_canvas.bind('<Configure>', lambda e: main_canvas.configure(scrollregion=main_canvas.bbox('all')))
main_canvas.bind_all('<MouseWheel>', lambda e: on_mousewheel(e))

minor_frame = Frame(main_canvas)
main_canvas.create_window((0,0), window=minor_frame, anchor='nw')

for i, text in enumerate(label_texts, start=1):
    group_frame = Frame(minor_frame, height=20, width=600)
    if (i == 1) or (i ==2):
        group_frame.grid(column=1, row=2*i-1, columnspan=3, padx=10, pady=10, sticky='we')
    else:
        group_frame.grid(column=4, row=2*i-5, columnspan=3, padx=10, pady=10, sticky='we')

    lb = Label(group_frame, text=f'Soubor {text}')
    lb.grid(column=1, row=1, columnspan=2, padx=10, sticky= 'w')

    t = Text(group_frame, height=1, width=30, background='#E0E0E0')
    t.insert('1.0', f'{d.basenames[i-1]}')
    t.configure(state='disabled')
    t.grid(column=2, row=2, padx=10, pady=5, sticky='w')

    b = Button(group_frame, text='Najít soubor', command=partial(browse_files, t, i-1, True))
    b.grid(column=1, row=2, padx=10, pady=5, sticky='w')

group_frame = Frame(minor_frame, height=20, width=600)
group_frame.grid(column=1, row=2*(i+1)-1, columnspan=3, padx=10, pady=10, sticky='we')

lb = Label(group_frame, text=f'Složka kam vložit výstup.')
lb.grid(column=1, row=1, columnspan=2, padx=10, sticky='w')

t = Text(group_frame, height=1, width=30, background='#E0E0E0')
t.insert('1.0', f'.\{os.sep.join(os.path.normpath(d.outputdir).split(os.sep)[-3:])}')
t.configure(state='disabled')
t.grid(column=2, row=2, padx=10, pady=5)

b = Button(group_frame, text='Vložit cestu', command=partial(browse_files, t, -1, False))
b.grid(column=1, row=2, padx=12, pady=5)

b_run = Button(minor_frame, text='Vytvořit přehled', command=run_create_table, background='#CACAFF')
b_run.grid(column=5, row=2*(i-1)+3)

labelos = Label(minor_frame, text='')
labelos.grid(column=5, row=2*(i-1)+2)

r.bind('<Return>', lambda e=None: press_create(b_run))

r.mainloop()
