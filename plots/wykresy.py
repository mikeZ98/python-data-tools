import pandas as pd
import matplotlib.pyplot as plt
import plotly.graph_objs as go
import plotly.io as pio
import tkinter as tk
from tkinter import filedialog, messagebox, StringVar, IntVar, Checkbutton, Frame, Button, Radiobutton, DISABLED, NORMAL
import os
from simpledbf import Dbf5
from datetime import datetime

# Funkcja do sprawdzenia obecności kolumny czasu we wszystkich plikach
def check_if_all_files_have_time_column(files_columns):
    all_have_time = True
    only_dbf_files = all(file.endswith('.dbf') for file, _ in files_columns)
    
    for _, df in files_columns:
        if 'pm_time' not in df.columns:
            all_have_time = False
            break
    
    return all_have_time and only_dbf_files

# Funkcja do synchronizacji plików DBF na podstawie kolumny z czasem 'pm_time'
def synchronize_dbf_data(files_columns):
    min_time = None
    max_time = None
    synced_dataframes = []
    time_col_found = False

    for file, df in files_columns:
        pm_time_columns = [col for col in df.columns if 'pm_time' in col]
        if len(pm_time_columns) == 0:
            continue

        time_col_found = True

        try:
            df[pm_time_columns[0]] = pd.to_datetime(df[pm_time_columns[0]], format='%Y-%m-%d %H:%M:%S', errors='coerce')
        except Exception as e:
            print(f"Błąd przy konwersji kolumny 'pm_time' w pliku {file}: {e}")
            continue
        
        if min_time is None or df[pm_time_columns[0]].min() < min_time:
            min_time = df[pm_time_columns[0]].min()
        if max_time is None or df[pm_time_columns[0]].max() > max_time:
            max_time = df[pm_time_columns[0]].max()

        synced_dataframes.append(df)

    if not time_col_found:
        return None, None, None

    for df in synced_dataframes:
        pm_time_columns = [col for col in df.columns if 'pm_time' in col]
        if len(pm_time_columns) > 0:
            df['TimeDiff'] = (df[pm_time_columns[0]] - min_time).dt.total_seconds()

    merged_df = pd.concat(synced_dataframes, axis=1) if synced_dataframes else None
    
    return merged_df, min_time, max_time

# Funkcja do dodania prefiksów do nazw kolumn w DataFrame na podstawie nazwy pliku
def add_prefix_to_columns(files_columns):
    prefixed_dataframes = []
    for file, df in files_columns:
        prefix = os.path.splitext(os.path.basename(file))[0]
        df.columns = [f"{prefix}_{col}" for col in df.columns]
        prefixed_dataframes.append(df)
    return prefixed_dataframes

# Funkcja do wybierania plików z różnych folderów
def select_files_from_different_folders():
    selected_files = []

    def add_files():
        files = filedialog.askopenfilenames(
            title="Wybierz pliki do połączenia", 
            filetypes=[
                ("Pliki CSV", "*.csv"), 
                ("Pliki Excel", "*.xls;*.xlsx"), 
                ("Pliki DBF", "*.dbf"),
                ("Wszystkie obsługiwane pliki", "*.csv;*.xls;*.xlsx;*.dbf")
            ],
            initialdir=os.getcwd()
        )
        if files:
            selected_files.extend(files)

    def finish_selection():
        if selected_files:
            root.destroy()
        else:
            messagebox.showwarning("Ostrzeżenie", "Nie wybrano żadnych plików. Proces został przerwany.")
    
    root = tk.Toplevel()
    root.title("Wybieranie plików z różnych folderów")

    label = tk.Label(root, text="Wybierz pliki z różnych folderów")
    label.pack()

    add_button = Button(root, text="Dodaj pliki", command=add_files)
    add_button.pack()

    finish_button = Button(root, text="Zakończ wybieranie", command=finish_selection)
    finish_button.pack()

    root.wait_window()

    return [(file, read_file(file)) for file in selected_files] if selected_files else None

def read_file(file):
    if file.endswith('.csv'):
        return pd.read_csv(file)
    elif file.endswith(('.xls', '.xlsx')):
        return pd.read_excel(file)
    elif file.endswith('.dbf'):
        dbf = Dbf5(file)
        return dbf.to_dataframe()

# Funkcja wyboru formatu wyjściowego
def select_output_format():
    format_window = tk.Toplevel()
    format_window.title("Wybierz format wyjściowy")

    output_format = tk.StringVar(value='csv')

    def submit_format():
        format_window.destroy()

    label = tk.Label(format_window, text="Wybierz format pliku wyjściowego:")
    label.pack()

    csv_check = Radiobutton(format_window, text="CSV", variable=output_format, value='csv')
    csv_check.pack(anchor=tk.W)

    xlsx_check = Radiobutton(format_window, text="Excel (xlsx)", variable=output_format, value='xlsx')
    xlsx_check.pack(anchor=tk.W)

    submit_button = Button(format_window, text="Zatwierdź", command=submit_format)
    submit_button.pack()

    format_window.wait_window()
    return output_format.get()

# Nowa funkcja do wyboru kolumn z każdego pliku
def select_columns_for_file(file, df):
    columns = df.columns
    selected_columns = []

    def select_all():
        for var in column_vars:
            var.set(1)

    def deselect_all():
        for var in column_vars:
            var.set(0)

    def submit_selection():
        nonlocal selected_columns
        selected_columns = [columns[i] for i, var in enumerate(column_vars) if var.get()]
        if not selected_columns:
            messagebox.showwarning("Ostrzeżenie", "Wybierz przynajmniej jedną kolumnę.")
            return
        window.destroy()

    window = tk.Toplevel()
    window.title(f"Wybierz kolumny do połączenia: {os.path.basename(file)}")

    column_vars = []
    for column in columns:
        var = IntVar()
        check = Checkbutton(window, text=column, variable=var)
        check.pack(anchor=tk.W)
        column_vars.append(var)

    select_all_button = Button(window, text="Zaznacz wszystkie", command=select_all)
    select_all_button.pack(pady=(10, 0))

    deselect_all_button = Button(window, text="Odznacz wszystkie", command=deselect_all)
    deselect_all_button.pack()

    submit_button = Button(window, text="Zatwierdź", command=submit_selection)
    submit_button.pack(pady=(10, 20))

    window.wait_window()

    return selected_columns

def generate_filename(prefix, extension, output_dir):
    current_time = datetime.now().strftime("%Y%m%d_%H%M%S")
    return os.path.join(output_dir, f"{prefix}_{current_time}.{extension}")

def generate_plots(df, output_dir, files_columns):
    columns = df.columns

    def select_columns_for_plot(enable_time_checkbox):
        plot_window = tk.Toplevel()
        plot_window.title("Wybierz kolumny dla osi X i Y")

        selected_x_column = StringVar()
        selected_y_columns = []
        use_time_as_x = IntVar()

        def on_x_checkbox_selected(col, var):
            # Zablokuj wszystkie inne checkboxy dla osi X po wybraniu jednej kolumny
            if var.get():
                selected_x_column.set(col)
                for i, other_var in enumerate(x_checkbox_vars):
                    if x_checkboxes[i]["text"] != col:
                        other_var.set(0)
                        x_checkboxes[i].config(state=DISABLED)
            else:
                selected_x_column.set('')
                for checkbox in x_checkboxes:
                    checkbox.config(state=NORMAL)

        def submit_plot_columns():
            nonlocal selected_y_columns
            if use_time_as_x.get():
                selected_x_column.set('pm_time')
            selected_y_columns = [columns[i] for i, var in enumerate(y_checkbox_vars) if var.get()]
            if not selected_x_column.get() or not selected_y_columns:
                messagebox.showwarning("Ostrzeżenie", "Wybierz kolumnę dla osi X i przynajmniej jedną kolumnę dla osi Y.")
                return
            plot_window.destroy()

        # Kolumny dla osi X
        x_frame = Frame(plot_window)
        x_frame.grid(row=0, column=0, padx=10, pady=10, sticky='n')
        tk.Label(x_frame, text="Kolumna dla osi X").pack(anchor=tk.W)

        x_checkbox_vars = []
        x_checkboxes = []
        for column in columns:
            var = IntVar()
            check = Checkbutton(x_frame, text=column, variable=var, command=lambda col=column, v=var: on_x_checkbox_selected(col, v))
            check.pack(anchor=tk.W)
            x_checkbox_vars.append(var)
            x_checkboxes.append(check)

        # Kolumny dla osi Y
        y_frame = Frame(plot_window)
        y_frame.grid(row=0, column=1, padx=10, pady=10, sticky='n')
        tk.Label(y_frame, text="Kolumna dla osi Y").pack(anchor=tk.W)

        y_checkbox_vars = []
        for column in columns:
            var = IntVar()
            check = Checkbutton(y_frame, text=column, variable=var)
            check.pack(anchor=tk.W)
            y_checkbox_vars.append(var)

        # Opcje dodatkowe
        options_frame = Frame(plot_window)
        options_frame.grid(row=0, column=2, padx=10, pady=10, sticky='n')

        time_checkbox = Checkbutton(options_frame, text="Użyj czasu jako osi X", variable=use_time_as_x)
        time_checkbox.pack(anchor=tk.W)
        if not enable_time_checkbox:
            time_checkbox.config(state=DISABLED)

        submit_button = Button(options_frame, text="Zatwierdź", command=submit_plot_columns)
        submit_button.pack(pady=20)

        plot_window.wait_window()
        return selected_x_column.get(), selected_y_columns

    enable_time_checkbox = check_if_all_files_have_time_column(files_columns)
    while True:
        generate_more = messagebox.askyesno("Generowanie wykresu", "Czy chcesz wygenerować wykres?")
        if not generate_more:
            break

        x_col, y_cols = select_columns_for_plot(enable_time_checkbox)

        if x_col and y_cols:
                fig = go.Figure()
                for y_col in y_cols:
                    fig.add_trace(go.Scatter(x=df[x_col], y=df[y_col], mode='lines', name=y_col))
                fig.update_layout(title='Wykres dynamiczny', xaxis_title=x_col, yaxis_title="Wartość")
                plot_path_html = os.path.join(output_dir, "wykres.html")
                pio.write_html(fig, file=plot_path_html, auto_open=False)
                messagebox.showinfo("Informacja", "Wykres został pomyślnie wygenerowany.")  

def main():
    root = tk.Tk()
    root.withdraw()

    output_dir = filedialog.askdirectory(title="Wybierz folder do zapisu")
    if not output_dir:
        messagebox.showerror("Błąd", "Nie wybrano folderu. Proces został przerwany.")
        return

    files_columns = select_files_from_different_folders()
    if files_columns is None:
        return

    # Wybór kolumn dla każdego pliku
    for i, (file, df) in enumerate(files_columns):
        selected_columns = select_columns_for_file(file, df)
        files_columns[i] = (file, df[selected_columns])

    if check_if_all_files_have_time_column(files_columns):
        merged_df, min_time, max_time = synchronize_dbf_data(files_columns)
    else:
        dataframes = [df for _, df in files_columns]
        merged_df = pd.concat(dataframes, axis=1)

    prefixed_dataframes = add_prefix_to_columns(files_columns)
    final_df = pd.concat(prefixed_dataframes, axis=1)

    generate_plots(final_df, output_dir, files_columns)
    print("Proces zakończony.")

if __name__ == "__main__":
    main()
