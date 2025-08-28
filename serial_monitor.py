import serial
import serial.tools.list_ports
import time
import datetime
import os

def select_port():
    """
    Lists available serial ports and prompts the user to select one.
    Returns the selected port name (e.g., 'COM3') or None if no ports are found.
    """
    print("
🔎 Searching for available serial ports...")
    ports = serial.tools.list_ports.comports() # Get a list of all available ports

    if not ports:
        print("🤔 No serial ports found. Please connect your device and try again.")
        return None

    print("📄 Available serial ports:")
    for i, port in enumerate(ports):
        # Display port info to help the user choose
        print(f"  [{i+1}] {port.device}: {port.description} [{port.hwid}]")

    while True:
        try:
            choice = int(input("
>> Please select a port (enter the number): "))
            if 1 <= choice <= len(ports):
                return ports[choice - 1].device
            else:
                print("⚠️ Invalid number. Please try again.")
        except (ValueError, IndexError):
            print("⚠️ Invalid input. Please enter a number from the list.")

def start_monitor(port, baudrate, log_dir):
    """
    Main function to monitor the serial port, handle autoreconnect,
    and write logs to a file.
    """
    print(f"
🚀 Starting serial monitor on {port} at {baudrate} baud.")
    print(f"✍️  Logging to directory: {log_dir}")
    print("Press Ctrl+C to stop the monitor.")

    # Create the log directory if it doesn't exist
    os.makedirs(log_dir, exist_ok=True)
    
    # Generate a unique log file name with a timestamp
    timestamp = datetime.datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    log_filename = os.path.join(log_dir, f"serial_log_{port}_{timestamp}.txt")
    
    print(f"📜 Log file: {log_filename}")

    current_serial = None

    try:
        while True:
            try:
                # If not connected, try to connect
                if current_serial is None or not current_serial.is_open:
                    print(f"🔌 Trying to connect to {port}...")
                    current_serial = serial.Serial(port, baudrate, timeout=1)
                    print(f"✅ Connection successful on {port}!")

                # Read a line from the serial port
                line = current_serial.readline()

                if line:
                    # Decode bytes to a string, replacing any problematic characters
                    decoded_line = line.decode('utf-8', errors='replace').strip()
                    
                    # Get current timestamp for the log entry
                    log_timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S.%f")[:-3]
                    
                    log_entry = f"{log_timestamp} | {decoded_line}"
                    
                    # 1. Output to console
                    print(log_entry)
                    
                    # 2. Write to log file
                    with open(log_filename, 'a', encoding='utf-8') as log_file:
                        log_file.write(log_entry + '\n')

            except serial.SerialException as e:
                # This exception is raised on disconnection or if port is not found
                if current_serial and current_serial.is_open:
                    current_serial.close()
                print(f"
❌ Port {port} disconnected or not available. {e}")
                print("🔄 Will attempt to reconnect automatically in 5 seconds...")
                time.sleep(5)
            except Exception as e:
                print(f"An unexpected error occurred: {e}")
                time.sleep(5)

    except KeyboardInterrupt:
        # Handle Ctrl+C to exit gracefully
        print("\n🛑 Monitor stopped by user.")
        if current_serial and current_serial.is_open:
            current_serial.close()
            print(f"🔌 Port {port} closed.")

if __name__ == "__main__":
    # --- 1. Port Selection ---
    selected_port = select_port()
    
    if selected_port:
        # --- 2. Baud Rate Selection ---
        common_baud_rates = [9600, 19200, 38400, 57600, 115200]
        while True:
            try:
                baud = int(input(f"
>> Enter baud rate (e.g., {', '.join(map(str, common_baud_rates))}): "))
                break
            except ValueError:
                print("⚠️ Invalid input. Please enter a whole number.")

        # --- 3. Log Directory Selection ---
        default_log_dir = os.path.join(os.path.expanduser("~"), "SerialLogs")
        log_directory = input(f"
>> Enter directory to save logs (press Enter for default: {default_log_dir}): ")
        if not log_directory:
            log_directory = default_log_dir

        # --- 4. Start the Monitor ---
        start_monitor(selected_port, baud, log_directory)