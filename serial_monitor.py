import tkinter as tk
from tkinter import ttk, filedialog, scrolledtext
import serial
import serial.tools.list_ports
import threading
import queue
import json
import os
import datetime
import time

# --- Configuration File ---
CONFIG_FILE = "serial_monitor_config.json"

class SerialMonitorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("GEMINI Serial Monitor 📈")
        self.root.geometry("800x600")

        # --- Member Variables ---
        self.serial_thread = None
        self.serial_port = None
        self.is_monitoring = False
        self.log_file_path = None
        self.serial_queue = queue.Queue()

        # --- Load saved settings ---
        self.settings = self.load_settings()

        # --- Build the GUI ---
        self.create_widgets()
        self.populate_ports()
        self.apply_settings()
        
        # --- Start the queue processor ---
        self.process_serial_queue()
        
        # --- Handle window closing ---
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

    def load_settings(self):
        """Loads settings from the JSON config file."""
        try:
            with open(CONFIG_FILE, 'r') as f:
                return json.load(f)
        except FileNotFoundError:
            return {"port": "", "baudrate": "921600", "log_dir": os.path.expanduser("~")}

    def save_settings(self):
        """Saves current settings to the JSON config file."""
        current_settings = {
            "port": self.port_var.get(),
            "baudrate": self.baud_var.get(),
            "log_dir": self.log_dir_var.get()
        }
        with open(CONFIG_FILE, 'w') as f:
            json.dump(current_settings, f, indent=4)

    def apply_settings(self):
        """Applies loaded settings to the GUI widgets."""
        self.port_var.set(self.settings.get("port", ""))
        self.baud_var.set(self.settings.get("baudrate", "921600"))
        self.log_dir_var.set(self.settings.get("log_dir", os.path.expanduser("~")))

    def create_widgets(self):
        """Creates and lays out all the GUI widgets."""
        control_frame = ttk.Frame(self.root, padding="10")
        control_frame.pack(side=tk.TOP, fill=tk.X)

        ttk.Label(control_frame, text="Port:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.port_var = tk.StringVar()
        self.port_combo = ttk.Combobox(control_frame, textvariable=self.port_var, width=30)
        self.port_combo.grid(row=0, column=1, padx=5, pady=5)
        ttk.Button(control_frame, text="Refresh", command=self.populate_ports).grid(row=0, column=2, padx=5, pady=5)

        ttk.Label(control_frame, text="Baud Rate:").grid(row=1, column=0, padx=5, pady=5, sticky="w")
        self.baud_var = tk.StringVar()
        
        # ⭐ CHANGED: Added 250000 to the list of baud rates
        baud_rates = ["9600", "19200", "38400", "57600", "115200", "230400", "250000", "460800", "921600"]
        
        self.baud_combo = ttk.Combobox(self.root, textvariable=self.baud_var, values=baud_rates, width=15)
        self.baud_combo.grid(row=1, column=1, in_=control_frame, padx=5, pady=5)

        ttk.Label(control_frame, text="Log Directory:").grid(row=2, column=0, padx=5, pady=5, sticky="w")
        self.log_dir_var = tk.StringVar()
        ttk.Entry(control_frame, textvariable=self.log_dir_var, width=50).grid(row=2, column=1, columnspan=2, padx=5, pady=5, sticky="we")
        ttk.Button(control_frame, text="Browse...", command=self.select_log_directory).grid(row=2, column=3, padx=5, pady=5)
        
        button_frame = ttk.Frame(self.root, padding="5")
        button_frame.pack(side=tk.TOP, fill=tk.X)
        self.start_button = ttk.Button(button_frame, text="▶ Start Monitoring", command=self.start_monitoring)
        self.start_button.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=10)
        self.stop_button = ttk.Button(button_frame, text="■ Stop Monitoring", command=self.stop_monitoring, state=tk.DISABLED)
        self.stop_button.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=10)

        self.output_text = scrolledtext.ScrolledText(self.root, wrap=tk.WORD, state=tk.DISABLED, bg="#2b2b2b", fg="#d3d3d3")
        self.output_text.pack(side=tk.TOP, expand=True, fill=tk.BOTH, padx=10, pady=5)

        self.status_var = tk.StringVar(value="Ready. Select a port and click Start.")
        status_bar = ttk.Label(self.root, textvariable=self.status_var, relief=tk.SUNKEN, anchor=tk.W, padding="2 5")
        status_bar.pack(side=tk.BOTTOM, fill=tk.X)

    def populate_ports(self):
        ports = [port.device for port in serial.tools.list_ports.comports()]
        self.port_combo['values'] = ports or ["No ports found"]
        if ports:
            saved_port = self.settings.get("port")
            if saved_port in ports:
                self.port_var.set(saved_port)
            else:
                self.port_var.set(ports[0])

    def select_log_directory(self):
        directory = filedialog.askdirectory(initialdir=self.log_dir_var.get())
        if directory:
            self.log_dir_var.set(directory)

    def start_monitoring(self):
        port = self.port_var.get()
        baudrate = self.baud_var.get()
        if not port or "No ports found" in port or not baudrate:
            self.status_var.set("Error: Port and Baud Rate must be selected.")
            return

        self.is_monitoring = True
        self.start_button.config(state=tk.DISABLED)
        self.stop_button.config(state=tk.NORMAL)
        self.port_combo.config(state=tk.DISABLED)
        self.baud_combo.config(state=tk.DISABLED)

        log_dir = self.log_dir_var.get()
        os.makedirs(log_dir, exist_ok=True)
        timestamp = datetime.datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        self.log_file_path = os.path.join(log_dir, f"serial_log_{timestamp}.txt")

        self.serial_thread = threading.Thread(target=self.serial_reader_thread, daemon=True)
        self.serial_thread.start()
        self.status_var.set(f"Connecting to {port}...")

    def stop_monitoring(self):
        self.is_monitoring = False
        if self.serial_thread and self.serial_thread.is_alive():
            self.serial_thread.join(timeout=1)

        self.start_button.config(state=tk.NORMAL)
        self.stop_button.config(state=tk.DISABLED)
        self.port_combo.config(state=tk.NORMAL)
        self.baud_combo.config(state=tk.NORMAL)
        self.status_var.set("Stopped. Ready to start again.")

    def serial_reader_thread(self):
        port = self.port_var.get()
        baudrate = int(self.baud_var.get())

        while self.is_monitoring:
            try:
                with serial.Serial(port, baudrate, timeout=1) as ser:
                    self.serial_queue.put(f"STATUS: ✅ Connected to {port} at {baudrate} baud.")
                    while self.is_monitoring:
                        try:
                            line = ser.readline()
                            if line:
                                self.serial_queue.put(line)
                        except serial.SerialException:
                            break
            except serial.SerialException:
                self.serial_queue.put(f"STATUS: ❌ Port {port} unavailable. Retrying in 5s...")
                time.sleep(5)

        if self.serial_port and self.serial_port.is_open:
            self.serial_port.close()
        self.serial_queue.put(f"STATUS: Disconnected from {port}.")

    def process_serial_queue(self):
        try:
            while not self.serial_queue.empty():
                message = self.serial_queue.get_nowait()
                if isinstance(message, str) and message.startswith("STATUS:"):
                    self.status_var.set(message.replace("STATUS: ", ""))
                elif isinstance(message, bytes):
                    try:
                        decoded_line = message.decode('utf-8').strip()
                        timestamp = datetime.datetime.now().strftime("%H:%M:%S.%f")[:-3]
                        full_line = f"[{timestamp}] {decoded_line}\n"
                        
                        self.output_text.config(state=tk.NORMAL)
                        self.output_text.insert(tk.END, full_line)
                        self.output_text.see(tk.END)
                        self.output_text.config(state=tk.DISABLED)

                        if self.log_file_path:
                            with open(self.log_file_path, 'a', encoding='utf-8') as f:
                                f.write(full_line)

                    except UnicodeDecodeError:
                        pass
        finally:
            self.root.after(100, self.process_serial_queue)

    def on_closing(self):
        self.stop_monitoring()
        self.save_settings()
        self.root.destroy()

if __name__ == "__main__":
    root = tk.Tk()
    app = SerialMonitorApp(root)
    root.mainloop()