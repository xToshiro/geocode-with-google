import tkinter as tk
from tkinter import filedialog, messagebox, ttk, scrolledtext
import pandas as pd
import threading
from selenium import webdriver
from selenium.webdriver.edge.service import Service
from webdriver_manager.microsoft import EdgeChromiumDriverManager
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import json
import logging
import os

# Setup logging
LOG_FILE = "geocoding.log"
CACHE_FILE = "geocoding_cache.json"
logging.basicConfig(filename=LOG_FILE, level=logging.INFO, format='%(asctime)s:%(levelname)s:%(message)s')

# Global flag to stop geocoding
stop_geocoding = False

def setup_selenium_edge():
    options = webdriver.EdgeOptions()
    options.add_argument("--headless")
    service = Service(EdgeChromiumDriverManager().install())
    driver = webdriver.Edge(service=service, options=options)
    logging.info("WebDriver session started.")
    return driver

def load_or_create_cache():
    if not os.path.isfile(CACHE_FILE):
        with open(CACHE_FILE, 'w') as file:
            json.dump({}, file)
        logging.info("New cache file created.")
    try:
        with open(CACHE_FILE, 'r') as file:
            cache = json.load(file)
    except json.JSONDecodeError:
        logging.error("Cache file is corrupted. Creating a new one.")
        cache = {}
    return cache

def update_cache(cache):
    with open(CACHE_FILE, 'w') as file:
        json.dump(cache, file, indent=4)

def get_lat_long(driver, address, cache, max_retries=3):
    if address in cache:
        return cache[address]['latitude'], cache[address]['longitude'], True
    retries = 0
    while retries < max_retries:
        try:
            driver.get(f"https://www.google.com/maps/search/{'+'.join(address.split())}")
            wait = WebDriverWait(driver, 10)
            wait.until(EC.url_contains("@"))
            url = driver.current_url
            coords = url.split('@')[1].split(',')[0:2]
            latitude, longitude = float(coords[0]), float(coords[1].split('!')[0])
            cache[address] = {'latitude': latitude, 'longitude': longitude}
            update_cache(cache)
            return latitude, longitude, True
        except Exception as e:
            logging.error(f"Retry {retries + 1} for address '{address}': {e}")
            retries += 1
    return None, None, False

def geocode_addresses(input_file, selected_columns, progress_callback, finished_callback, status_callback):
    global stop_geocoding
    status_callback("Preparing data... This may take some time...")
    df_input = pd.read_excel(input_file)
    
    driver = setup_selenium_edge()
    cache = load_or_create_cache()
    
    status_callback("Geocoding data...")
    for index, row in df_input.iterrows():
        if stop_geocoding:
            break
        address_parts = [str(row[col]).strip() for col in selected_columns if col in df_input.columns]
        address = ' '.join(address_parts)
        latitude, longitude, success = get_lat_long(driver, address, cache)
        if success:
            df_input.at[index, 'Latitude'] = latitude
            df_input.at[index, 'Longitude'] = longitude
            progress_callback(index + 1, df_input.shape[0])
    
    update_cache(cache)
    driver.quit()
    
    output_file = os.path.splitext(input_file)[0] + "_geocoded.xlsx"
    df_input.to_excel(output_file, index=False)
    finished_callback(output_file, not stop_geocoding)

class GeocodeApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Transporte e Meio Ambiente(Trama) - Geocoding with Google - App")
        self.geometry("600x490")  # Increased height for better spacing
        
        self.create_widgets()
        self.file_path = ""
        self.column_checkboxes = []
        self.selected_columns = []
        
    def create_widgets(self):
        self.file_label = tk.Label(self, text="No file selected")
        self.file_label.pack(pady=10)
        
        self.browse_button = tk.Button(self, text="Browse", command=self.browse_file)
        self.browse_button.pack()

        self.columns_frame = tk.LabelFrame(self, text="Columns")
        self.columns_scroll = scrolledtext.ScrolledText(self.columns_frame, width=40, height=10, wrap=tk.WORD)
        self.columns_scroll.pack()
        self.columns_frame.pack(pady=10, fill="both", expand="yes")
        
        self.status_label = tk.Label(self, text="")
        self.status_label.pack(pady=10)

        self.start_button = tk.Button(self, text="Start Geocoding", command=self.start_geocoding)
        self.start_button.pack(pady=5)
        
        self.stop_button = tk.Button(self, text="Stop Geocoding", command=self.stop_geocoding, state=tk.DISABLED)
        self.stop_button.pack(pady=5)

        self.progress_label = tk.Label(self, text="Progress: 0/0")
        self.progress_label.pack(pady=5)
        
        self.progress = ttk.Progressbar(self, orient=tk.HORIZONTAL, length=300, mode='determinate')
        self.progress.pack(pady=20)  # Added padding for spacing from the bottom

    def browse_file(self):
        global stop_geocoding
        stop_geocoding = False
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if file_path:
            self.file_path = file_path
            self.file_label.config(text=f"File: {os.path.basename(file_path)}")
            self.display_column_checkboxes(file_path)

    def display_column_checkboxes(self, file_path):
        for widget in self.columns_frame.winfo_children():
            widget.destroy()
        
        df = pd.read_excel(file_path)
        self.selected_columns = []
        for col in df.columns:
            var = tk.BooleanVar()
            chk = tk.Checkbutton(self.columns_frame, text=col, variable=var)
            chk.pack(anchor='w')
            self.column_checkboxes.append((var, col))
    
    def update_progress(self, current, total):
        self.progress['value'] = (current / total) * 100
        self.progress_label.config(text=f"Progress: {current}/{total}")
        self.update_idletasks()

    def finished_geocoding(self, output_file, success):
        if success:
            messagebox.showinfo("Complete", f"Geocoding complete. Output saved to {output_file}")
        else:
            messagebox.showinfo("Stopped", "Geocoding stopped by user.")
        self.start_button['state'] = tk.NORMAL
        self.stop_button['state'] = tk.DISABLED
    
    def update_status(self, message):
        self.status_label.config(text=message)
        self.update_idletasks()

    def start_geocoding(self):
        global stop_geocoding
        stop_geocoding = False
        selected_columns = [col for var, col in self.column_checkboxes if var.get()]
        if not self.file_path or not selected_columns:
            messagebox.showerror("Error", "Please select a file and at least one column.")
            return
        self.start_button['state'] = tk.DISABLED
        self.stop_button['state'] = tk.NORMAL
        threading.Thread(target=geocode_addresses, args=(self.file_path, selected_columns, self.update_progress, self.finished_geocoding, self.update_status)).start()

    def stop_geocoding(self):
        global stop_geocoding
        stop_geocoding = True
        self.stop_button['state'] = tk.DISABLED

if __name__ == "__main__":
    app = GeocodeApp()
    app.mainloop()
