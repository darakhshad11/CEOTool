import threading

import pandas as pd
import random
import requests
from bs4 import BeautifulSoup
import time
import tkinter as tk
from tkinter import filedialog


def get_ceo_name(company_name, sinr):
    import threading
    url = "https://www.google.com/search?q=" + company_name +" "+ sinr
    webpage = requests.get(url)

    if "Our systems have detected unusual traffic from your computer" in webpage.text:
        print("ip blocked...")
        log.config(state='normal')
        log.insert(tk.END, f"ip blocked... program quiting")
        log.config(state='disable')
        exit()
        return "IP Blocked"
    soup = BeautifulSoup(webpage.text, 'html.parser')

    try:
        CeoName = soup.find("div", {"class": "BNeawe iBp4i AP7Wnd"}).text
    except:
        try:
            CeoName = soup.find("div", {"class": "kCrYT"}).span.text
        except:

            CeoName = "not found"
            log.config(state='normal')
            log.insert(tk.END, f"not found \n")
            log.config(state='disable')
    return CeoName

def main(file,sinroity,delay):
    df = pd.read_excel(file)
    df['Previous CEO Name'] = df['Ceo Name']
    #df["Ceo Name"] = ""
    start_row = int(start_row_input.get())
    for i in range(start_row-1,len(df)):
        try:
            company_name =  df.iloc[i, 0]
            ceo_name = get_ceo_name(company_name,sinroity)
            df.at[i, 'Ceo Name'] = ceo_name
            df.at[i, 'Previous CEO Name'] = df.at[i, 'Ceo Name']
            df.to_excel(file, index=False)
            log.config(state='normal')
            if ceo_name:
                log.insert(tk.END, f"Found {ceo_name} for {company_name} at row {i+1} of {len(df)}\n")
            else:
                log.insert(tk.END, f"CEO not found for {company_name} at row {i+1} of {len(df)}\n")
            log.config(state='disable')
            time.sleep(random.randint(1,int(delay)))
        except Exception as e:
            log.config(state='normal')
            log.insert(tk.END, f"An error occurred: {e}\n")
            log.config(state='disable')
            return "An error occurred"

def browse_file():
    global file
    filepath = filedialog.askopenfilename()
    file=filepath
    return filepath

def run_script():
    sinroity = sinroity_entry.get()
    delay = delay_entry.get()
    log.config(state='normal')
    thread = threading.Thread(target=main, args=(file,sinroity, delay))
    thread.start()
    log.insert(tk.END, f"Started scraping for {file}\n")
    log.config(state='disable')
def fun_quit():
    log.config(state='normal')
    log.insert(tk.END, f"scrapper bot exiting...\n")
    log.config(state='disable')
    root.destroy()
def add_placeholder(entry, text):
    entry.insert(0, text)
    entry.config(fg='grey')
    entry.bind("<FocusIn>", lambda args: entry.delete('0', 'end') if entry['fg'] == 'grey' else None)
    entry.bind("<FocusOut>", lambda args: entry.insert(0, text) if not entry.get() else None)
root = tk.Tk()
root.title("CEO Scraper")

browse_button = tk.Button(root, text="Browse", command=browse_file)
browse_button.grid(row=0, column=0, padx=5, pady=5)

sinroity_label = tk.Label(root, text="Sinroity")
sinroity_label.grid(row=1, column=0, padx=5, pady=5)

sinroity_entry = tk.Entry(root)
sinroity_entry.grid(row=1, column=1, padx=5, pady=5)
add_placeholder(sinroity_entry, "Enter the role")
start_row_label = tk.Label(root, text="Starting Row")
start_row_label.grid(row=2, column=0, padx=5, pady=5)

start_row_input = tk.Entry(root)
start_row_input.grid(row=2, column=1, padx=5, pady=5)
add_placeholder(start_row_input,"Starting Row Number")
delay_label = tk.Label(root, text="Delay (sec)")
delay_label.grid(row=4, column=0, padx=5, pady=5)

delay_entry = tk.Entry(root)
delay_entry.grid(row=4, column=1, padx=5, pady=5)
add_placeholder(delay_entry,"Time Range")
run_button = tk.Button(root, text="Run", command=run_script)

run_button.grid(row=6, column=0, padx=5, pady=5)

quit_button = tk.Button(root, text="Quit", command=fun_quit)
quit_button.grid(row=5, column=1, padx=5, pady=5)

log = tk.Text(root)
log.grid(row=5, column=0)
log.config(state='disable')

root.mainloop()
