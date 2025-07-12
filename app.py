import tkinter as tk
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
import pandas as pd
from datetime import datetime
import os
import json

# Määritä kansio, johon tiedostot tallennetaan. Muuta tämä haluamaasi polkuun.
output_folder = r"C:\path\to\folder"  # Esimerkiksi r"C:\Users\Isäsi\Dokumentit\Työ"

# Väliaikainen tiedosto merkinnöille
temp_file = "temp_entries.json"

# Luo pääikkuna modernilla teemalla
root = ttk.Window(themename="flatly")  # Moderni teema, voit kokeilla myös 'minty', 'vapor' jne.
root.title("Työseuranta")
root.geometry("900x700")  # Suurempi ikkunakoko paremmalle näkyvyydelle

# Aseta suurempi fontti koko sovellukselle (ttk-widgeteille)
style = ttk.Style()
style.configure('.', font=('Helvetica', 16))  # Suurempi fontti (16 pt)
style.configure('TButton', font=('Helvetica', 18))  # Vielä suurempi painikkeille
style.configure('TLabel', font=('Helvetica', 18))
style.configure('TEntry', font=('Helvetica', 18))

# Lista merkintöjen tallentamiseen
entries = []

# Muuttuja muokkaustilalle
editing_index = None

# Kuukausien lista
months = ["Tammikuu", "Helmikuu", "Maaliskuu", "Huhtikuu", "Toukokuu", "Kesäkuu", "Heinäkuu", "Elokuu", "Syyskuu", "Lokakuu", "Marraskuu", "Joulukuu"]

# Funktio merkintöjen lataamiseksi temp-tiedostosta
def load_entries():
    global entries
    if os.path.exists(temp_file):
        with open(temp_file, 'r') as f:
            entries = json.load(f)
        # Päivitä listbox
        for entry in entries:
            listbox.insert(tk.END, f"{entry['Päivämäärä']} | {entry['Työtunnit']} tuntia | {entry['Ajokilometrit']} km | {entry['Työpaikka']}")
        status_label.config(text="Edelliset merkinnät ladattu onnistuneesti!")

# Funktio merkintöjen tallentamiseksi temp-tiedostoon
def save_temp_entries():
    with open(temp_file, 'w') as f:
        json.dump(entries, f)

# Funktio merkinnän lisäämiseksi tai muokkaamiseksi
def save_entry():
    global editing_index
    try:
        date = date_entry.entry.get()
        hours = hours_entry.get()
        km = km_entry.get()
        place = place_entry.get().capitalize()  # Isoita ensimmäinen kirjain
        
        if not date or not hours or not km or not place:
            status_label.config(text="Täytä kaikki kentät ja valitse päivämäärä!")
            return
        
        # Tarkista päivämäärän muoto ja numerot
        datetime.strptime(date, '%d.%m.%Y')
        hours_val = float(hours)
        km_val = float(km)
        if hours_val < 0 or km_val < 0:
            status_label.config(text="Työtunnit ja kilometrit eivät voi olla negatiivisia!")
            return
        if hours_val > 24:
            status_label.config(text="Työtunnit eivät voi olla yli 24 tuntia!")
            return
        
        # Tarkista duplikaatti, mutta ohita jos muokataan samaa
        for idx, entry in enumerate(entries):
            if entry['Päivämäärä'] == date and idx != editing_index:
                status_label.config(text="Tämä päivämäärä on jo lisätty!")
                return
    except ValueError:
        status_label.config(text="Virheellinen syöte! Tunnit/Km: numerot")
        return
    
    entry_data = {
        'Päivämäärä': date,
        'Työtunnit': hours,
        'Ajokilometrit': km,
        'Työpaikka': place
    }
    
    if editing_index is not None:
        # Muokkaa olemassa olevaa
        entries[editing_index] = entry_data
        listbox.delete(editing_index)
        listbox.insert(editing_index, f"{date} | {hours} tuntia | {km} km | {place}")
        status_label.config(text="Merkintä muokattu onnistuneesti!")
        editing_index = None
        add_button.config(text="Lisää merkintä")
    else:
        # Lisää uusi
        entries.append(entry_data)
        listbox.insert(tk.END, f"{date} | {hours} tuntia | {km} km | {place}")
        status_label.config(text="Merkintä lisätty onnistuneesti!")
    
    # Tyhjennä kentät
    date_entry.entry.delete(0, END)
    hours_entry.delete(0, END)
    km_entry.delete(0, END)
    place_entry.delete(0, END)
    
    # Tallenna temp-tiedostoon
    save_temp_entries()

# Funktio muokkauksen aloittamiseksi
def edit_entry():
    global editing_index
    selected = listbox.curselection()
    if not selected:
        status_label.config(text="Valitse merkintä listasta muokataksesi!")
        return
    
    editing_index = selected[0]
    entry = entries[editing_index]
    
    date_entry.entry.delete(0, END)
    date_entry.entry.insert(0, entry['Päivämäärä'])
    hours_entry.delete(0, END)
    hours_entry.insert(0, entry['Työtunnit'])
    km_entry.delete(0, END)
    km_entry.insert(0, entry['Ajokilometrit'])
    place_entry.delete(0, END)
    place_entry.insert(0, entry['Työpaikka'])
    
    add_button.config(text="Tallenna muutokset")
    status_label.config(text="Muokkaa kenttiä ja tallenna.")

# Funktio vienniksi Exceliin
def export_to_excel():
    if not entries:
        status_label.config(text="Ei merkintöjä vietäväksi!")
        return
    
    # Hae kuukausi ja vuosi ensimmäisestä merkinnästä
    first_date = entries[0]['Päivämäärä']
    dt = datetime.strptime(first_date, '%d.%m.%Y')
    month_num = dt.month
    year = dt.year
    month = months[month_num - 1]  # Kuukausi suomeksi listasta
    
    # Luo vuosi-kansio, jos ei ole olemassa
    year_folder = os.path.join(output_folder, str(year))
    os.makedirs(year_folder, exist_ok=True)
    
    df = pd.DataFrame(entries)
    
    # Lajittele päivämäärän mukaan
    df['date_obj'] = df['Päivämäärä'].apply(lambda x: datetime.strptime(x, '%d.%m.%Y'))
    df = df.sort_values('date_obj')
    df = df.drop('date_obj', axis=1)
    
    # Laske summat
    total_hours = df['Työtunnit'].astype(float).sum()
    total_km = df['Ajokilometrit'].astype(float).sum()
    
    # Lisää yhteenvetorivi
    total_row = {
        'Päivämäärä': 'Yhteensä',
        'Työtunnit': total_hours,
        'Ajokilometrit': total_km,
        'Työpaikka': ''
    }
    total_df = pd.DataFrame([total_row])
    df = pd.concat([df, total_df], ignore_index=True)
    
    # Muodosta tiedostonimi ja polku
    file_name = f"{month} {year}.xlsx"
    file_path = os.path.join(year_folder, file_name)
    
    df.to_excel(file_path, index=False)
    status_label.config(text=f"Viety tiedostoon {file_path}")
    
    # Poista temp-tiedosto viennin jälkeen
    if os.path.exists(temp_file):
        os.remove(temp_file)
    entries.clear()
    listbox.delete(0, END)

# Pääkehys
frame = ttk.Frame(root, padding=30)  # Suurempi padding tilan lisäämiseksi
frame.grid(row=0, column=0, sticky='nsew')
frame.grid_rowconfigure(5, weight=1)
frame.grid_columnconfigure(0, weight=1)
frame.grid_columnconfigure(1, weight=2)  # Anna enemmän tilaa kentille

ttk.Label(frame, text="Päivämäärä:").grid(row=0, column=0, padx=15, pady=15, sticky='e')
date_entry = ttk.DateEntry(frame, dateformat="%d.%m.%Y", bootstyle=PRIMARY)
date_entry.grid(row=0, column=1, padx=15, pady=15, sticky='we')

ttk.Label(frame, text="Työtunnit:").grid(row=1, column=0, padx=15, pady=15, sticky='e')
hours_entry = ttk.Entry(frame)
hours_entry.grid(row=1, column=1, padx=15, pady=15, sticky='we')

ttk.Label(frame, text="Ajokilometrit:").grid(row=2, column=0, padx=15, pady=15, sticky='e')
km_entry = ttk.Entry(frame)
km_entry.grid(row=2, column=1, padx=15, pady=15, sticky='we')

ttk.Label(frame, text="Työpaikka:").grid(row=3, column=0, padx=15, pady=15, sticky='e')
place_entry = ttk.Entry(frame)
place_entry.grid(row=3, column=1, padx=15, pady=15, sticky='we')

# Lisää/Tallenna-painike
add_button = ttk.Button(frame, text="Lisää merkintä", command=save_entry, bootstyle=SUCCESS)
add_button.grid(row=4, column=0, pady=20, sticky='we')

# Muokkaa-painike
edit_button = ttk.Button(frame, text="Muokkaa merkintää", command=edit_entry, bootstyle=WARNING)
edit_button.grid(row=4, column=1, pady=20, sticky='we')

# Lista merkinnöistä
listbox = tk.Listbox(frame, height=12, font=('Helvetica', 16))  # Suurempi fontti suoraan listboxissa
listbox.grid(row=5, column=0, columnspan=2, padx=15, pady=15, sticky='nsew')

# Vie-painike
export_button = ttk.Button(frame, text="Vie Exceliin", command=export_to_excel, bootstyle=INFO)
export_button.grid(row=6, column=0, columnspan=2, pady=20, sticky='we')

# Tilaviesti
status_label = ttk.Label(frame, text="")
status_label.grid(row=7, column=0, columnspan=2, pady=15)

# Tee ikkunasta responsiivinen
root.grid_rowconfigure(0, weight=1)
root.grid_columnconfigure(0, weight=1)

# Lataa merkinnät käynnistyessä
load_entries()

# Käynnistä GUI
root.mainloop()