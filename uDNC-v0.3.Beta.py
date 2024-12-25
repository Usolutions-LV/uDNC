import tkinter as tk
from tkinter import filedialog
from tkinter import simpledialog
import json
import serial
import time
import threading
import os
import configparser
import webbrowser
import sys

# Default Settings
SETTINGS_FILE = 'settings.ini'
default_settings = {
    "COM_PORT": "COM1",
    "BAUD_RATE": 9600,
    "DATA_BITS": 8,
    "STOP_BITS": 2,
    "PARITY": "NONE",
    "FLOW_CONTROL": "Software",
    "LOGGING_ENABLED": False,
    "TRANSMISSION": True
}

# Load settings
config = configparser.ConfigParser()
if os.path.exists(SETTINGS_FILE):
    config.read(SETTINGS_FILE)
else:
    config["Settings"] = default_settings
    with open(SETTINGS_FILE, 'w') as f:
        config.write(f)

settings = {
    "COM_PORT":        config['Settings'].get("COM_PORT", default_settings["COM_PORT"]),
    "BAUD_RATE":       int(config['Settings'].get("BAUD_RATE", default_settings["BAUD_RATE"])),
    "DATA_BITS":       int(config['Settings'].get("DATA_BITS", default_settings["DATA_BITS"])),
    "STOP_BITS":       int(config['Settings'].get("STOP_BITS", default_settings["STOP_BITS"])),
    "PARITY":          config['Settings'].get("PARITY", default_settings["PARITY"]),
    "LOGGING_ENABLED": config["Settings"].getboolean("LOGGING_ENABLED", fallback=default_settings["LOGGING_ENABLED"]),
    "TRANSMISSION":    config["Settings"].getboolean("TRANSMISSION", fallback=default_settings["TRANSMISSION"]),
}

# Global Variables
COM_PORT        = settings["COM_PORT"]
BAUD_RATE       = settings["BAUD_RATE"]
DATA_BITS       = settings["DATA_BITS"]
STOP_BITS       = settings["STOP_BITS"]
PARITY          = settings["PARITY"]
LOGGING_ENABLED = settings["LOGGING_ENABLED"]
TRANSMISSION    = settings["TRANSMISSION"]

SEND_LOG_FILE    = 'datalog_send.log'
RECEIVE_LOG_FILE = 'datalog_receive.log'
FILENAME         = None
OUTPUT_FILENAME  = None
CYCLE_SEND       = False
STOP_REQUESTED   = False
ACTIVE_PROCESS   = None

# Function to save settings
def save_settings():
    config_to_save = configparser.ConfigParser()
    config_to_save["Settings"] = {
        "COM_PORT":        settings["COM_PORT"],
        "BAUD_RATE":       str(settings["BAUD_RATE"]),
        "DATA_BITS":       str(settings["DATA_BITS"]),
        "STOP_BITS":       str(settings["STOP_BITS"]),
        "PARITY":          settings["PARITY"],
        "FLOW_CONTROL":    settings["FLOW_CONTROL"],
        "LOGGING_ENABLED": str(settings["LOGGING_ENABLED"]),  # "True" or "False"
        "TRANSMISSION":    str(settings["TRANSMISSION"]),     # "True" or "False"
    }
    with open(SETTINGS_FILE, 'w') as f:
        config_to_save.write(f)

# Function to log data to a file
def log_data(file_path, data):
    if LOGGING_ENABLED:
        timestamp = time.strftime("%Y-%m-%d %H:%M:%S", time.gmtime())
        with open(file_path, 'a') as f:
            f.write(f"{timestamp} - {data}\n")

# Function to update the log in the GUI
def update_log(message):
    log_text.insert(tk.END, message + "\n")
    log_text.see(tk.END)

# Map string parity values to pyserial constants
PARITY_MAP = {
    "NONE":  serial.PARITY_NONE,
    "EVEN":  serial.PARITY_EVEN,
    "ODD":   serial.PARITY_ODD,
    "MARK":  serial.PARITY_MARK,
    "SPACE": serial.PARITY_SPACE,
}

# Function to force close the COM port if it's locked
def force_close_com(port_name):
    try:
        ser = serial.Serial(port_name)
        ser.close()
        update_log(f"Force-closed {port_name}.")
    except serial.SerialException as e:
        update_log(f"{port_name} not active or accessible: {e}")
    except Exception as e:
        update_log(f"Unexpected error while closing {port_name}: {e}")

# Function to send the .nc file
def send_file():
    global STOP_REQUESTED, CYCLE_SEND, ACTIVE_PROCESS
    ACTIVE_PROCESS = "send"
    STOP_REQUESTED = False

    # Ensure COM port is released
    force_close_com(COM_PORT)

    try:
        # Initialize the COM port
        ser = serial.Serial(
            port=COM_PORT,
            baudrate=BAUD_RATE,
            bytesize=DATA_BITS,
            stopbits=STOP_BITS,
            parity=PARITY_MAP.get(PARITY, serial.PARITY_NONE),
            timeout=1,
        )
        update_log(f"Opened {COM_PORT} successfully.")
    except serial.SerialException as e:
        update_log(f"SerialException while accessing {COM_PORT}: {e}")
        ACTIVE_PROCESS = None
        update_gui_buttons()
        return
    except PermissionError as e:
        update_log(f"PermissionError while accessing {COM_PORT}: {e}")
        ACTIVE_PROCESS = None
        update_gui_buttons()
        return
    except Exception as e:
        update_log(f"Unexpected error while accessing {COM_PORT}: {e}")
        ACTIVE_PROCESS = None
        update_gui_buttons()
        return

    # Start sending the file
    try:
        with open(FILENAME, 'r') as file:
            lines = file.readlines()
            update_log(f"Sending file: {FILENAME} with {len(lines)} lines.")
            RTS = False  # Ready-to-send flag

            # Wait for initial XON signal (if TRANSMISSION == True)
            update_log("Waiting for XON to start transmission...")
            while True:
                if TRANSMISSION is False:
                    update_log("No XON required by settings.")
                    RTS = True
                    break

                if ser.in_waiting > 0:
                    incoming_data = ser.read(ser.in_waiting)
                    if b'\x11' in incoming_data:  # XON
                        RTS = True
                        update_log("Received XON, starting transmission.")
                        break
                    elif b'\x13' in incoming_data:  # XOFF
                        update_log("Received XOFF, waiting for XON...")
                time.sleep(0.01)

            for idx, line in enumerate(lines, start=1):
                if STOP_REQUESTED:
                    update_log("Stop requested. Closing port and exiting send loop.")
                    ser.close()
                    ACTIVE_PROCESS = None
                    update_gui_buttons()
                    return

                # Check for XOFF during transmission
                while ser.in_waiting > 0:
                    incoming_data = ser.read(ser.in_waiting)
                    if b'\x13' in incoming_data:  # XOFF
                        RTS = False
                        update_log("Received XOFF, pausing transmission...")
                    elif b'\x11' in incoming_data:  # XON
                        RTS = True
                        update_log("Received XON, resuming transmission.")

                # Send line if RTS is True
                if RTS:
                    line = line.strip()  # Remove extra whitespace/newlines
                    if line:
                        data_to_send = line + '\r\n'
                        ser.write(data_to_send.encode('utf-8'))
                        update_log(f"Line {idx}: Sent: {data_to_send.strip()}")
                        log_data(SEND_LOG_FILE, f"Line {idx}: Sent: {data_to_send.strip()}")

                time.sleep(0.1)

        update_log("File sent successfully.")
        log_data(SEND_LOG_FILE, "File sent successfully.")
    except Exception as e:
        update_log(f"Error during file sending: {e}")
    finally:
        # Close the COM port at the end
        ser.close()
        update_log("COM port closed after sending.")
        ACTIVE_PROCESS = None
        update_gui_buttons()

def partial_send_file(start_line=None, start_index=None):
    """
    Sends the .nc file starting from a specific line number or Nxxx index, ensuring the first line is always sent
    and constraints like the second `%` are respected.

    :param start_line: Line number (1-based index) to start sending from.
    :param start_index: Nxxx index to start sending from (mutually exclusive with start_line).
    """
    global STOP_REQUESTED, ACTIVE_PROCESS
    ACTIVE_PROCESS = "partial_send"
    STOP_REQUESTED = False

    # Ensure COM port is released
    force_close_com(COM_PORT)

    try:
        # Initialize the COM port
        ser = serial.Serial(
            port=COM_PORT,
            baudrate=BAUD_RATE,
            bytesize=DATA_BITS,
            stopbits=STOP_BITS,
            parity=PARITY_MAP.get(PARITY, serial.PARITY_NONE),
            timeout=1,
        )
        update_log(f"Opened {COM_PORT} successfully.")
    except serial.SerialException as e:
        update_log(f"SerialException while accessing {COM_PORT}: {e}")
        ACTIVE_PROCESS = None
        update_gui_buttons()
        return
    except Exception as e:
        update_log(f"Unexpected error while accessing {COM_PORT}: {e}")
        ACTIVE_PROCESS = None
        update_gui_buttons()
        return

    try:
        with open(FILENAME, 'r') as file:
            lines = file.readlines()
            total_lines = len(lines)
            update_log(f"Loaded file: {FILENAME} with {total_lines} lines.")

            # Ensure the first line is always sent
            start_idx = 0  # Always include the first line (%)
            max_idx = total_lines - 1

            # Find the second `%` marker (if it exists)
            for i, line in enumerate(lines):
                if line.strip() == "%" and i > 0:
                    max_idx = i
                    break

            update_log(f"Maximum line allowed for transmission: {max_idx + 1} (before second `%`)")

            # Determine the starting point
            if start_line:
                start_idx = max(1, start_line - 1)  # Always start from at least the second line
            elif start_index:
                for i, line in enumerate(lines):
                    if line.lstrip().startswith(f'N{start_index}'):
                        start_idx = i
                        break
                else:
                    update_log(f"N{start_index} not found in the file. Cancelling operation.")
                    ser.close()
                    ACTIVE_PROCESS = None
                    update_gui_buttons()
                    return

            if start_idx > max_idx:
                update_log(f"Starting point ({start_idx + 1}) exceeds allowable range (up to {max_idx + 1}).")
                ser.close()
                ACTIVE_PROCESS = None
                update_gui_buttons()
                return

            # Log the actual starting point
            update_log(f"Starting transmission from line {start_idx + 1}.")

            # Ready-to-send flag
            RTS = False

            # Wait for initial XON signal if required
            if TRANSMISSION:
                update_log("Waiting for XON to start transmission...")
                while True:
                    if ser.in_waiting > 0:
                        incoming_data = ser.read(ser.in_waiting)
                        if b'\x11' in incoming_data:  # XON
                            RTS = True
                            update_log("Received XON, starting transmission.")
                            break
                        elif b'\x13' in incoming_data:  # XOFF
                            update_log("Received XOFF, waiting for XON...")
                    time.sleep(0.01)
            else:
                RTS = True

            # Transmit lines
            for idx, line in enumerate(lines[start_idx:max_idx + 1], start=start_idx + 1):
                if STOP_REQUESTED:
                    update_log("Stop requested. Closing port and exiting transmission loop.")
                    break

                # Handle XOFF/XON during transmission
                while ser.in_waiting > 0:
                    incoming_data = ser.read(ser.in_waiting)
                    if b'\x13' in incoming_data:  # XOFF
                        RTS = False
                        update_log("Received XOFF, pausing transmission...")
                    elif b'\x11' in incoming_data:  # XON
                        RTS = True
                        update_log("Received XON, resuming transmission.")

                if RTS:
                    line = line.strip()
                    if line:
                        data_to_send = line + '\r\n'
                        ser.write(data_to_send.encode('utf-8'))
                        update_log(f"Line {idx}: Sent: {data_to_send.strip()}")
                        log_data(SEND_LOG_FILE, f"Line {idx}: Sent: {data_to_send.strip()}")

                time.sleep(0.1)

        update_log("Partial file transmission completed.")
        log_data(SEND_LOG_FILE, "Partial file transmission completed.")
    except Exception as e:
        update_log(f"Error during partial file sending: {e}")
    finally:
        ser.close()
        update_log("COM port closed after partial sending.")
        ACTIVE_PROCESS = None
        update_gui_buttons()

# Function to receive the .nc file
def receive_file():
    global STOP_REQUESTED, ACTIVE_PROCESS
    ACTIVE_PROCESS = "receive"
    STOP_REQUESTED = False

    # Ensure COM port is released
    force_close_com(COM_PORT)

    try:
        # Initialize the COM port
        ser = serial.Serial(
            port=COM_PORT,
            baudrate=BAUD_RATE,
            bytesize=DATA_BITS,
            stopbits=STOP_BITS,
            parity=PARITY_MAP.get(PARITY, serial.PARITY_NONE),
            timeout=1,
        )
        update_log(f"Opened {COM_PORT} successfully.")
    except serial.SerialException as e:
        update_log(f"SerialException while accessing {COM_PORT}: {e}")
        ACTIVE_PROCESS = None
        update_gui_buttons()
        return
    except PermissionError as e:
        update_log(f"PermissionError while accessing {COM_PORT}: {e}")
        ACTIVE_PROCESS = None
        update_gui_buttons()
        return
    except Exception as e:
        update_log(f"Unexpected error while accessing {COM_PORT}: {e}")
        ACTIVE_PROCESS = None
        update_gui_buttons()
        return

    # Ask for the output file path
    output_path = filedialog.asksaveasfilename(defaultextension=".nc", filetypes=[["NC Files", "*.nc"]])
    if not output_path:
        update_log("No file selected for saving. Exiting receive loop.")
        ser.close()
        ACTIVE_PROCESS = None
        update_gui_buttons()
        return

    try:
        with open(output_path, 'w') as file:
            update_log(f"Receiving data and saving to {output_path}...")
            line_number   = 0
            percent_count = 0  # Tracks the number of '%' symbols received

            while True:
                if STOP_REQUESTED:
                    update_log("Stop requested. Closing port and exiting receive loop.")
                    ser.close()
                    ACTIVE_PROCESS = None
                    update_gui_buttons()
                    return

                # Read incoming data
                if ser.in_waiting > 0:
                    raw_data = ser.read(ser.in_waiting).decode('utf-8', errors='ignore')

                    # Remove all special characters except CR
                    sanitized_data = "".join(char if char == '\r' or char.isprintable() else '' for char in raw_data)

                    file.write(sanitized_data)
                    file.flush()

                    # Process each line split by \r and log it
                    for line in sanitized_data.split('\r'):
                        line_number += 1
                        stripped_line = line.strip()
                        if stripped_line:
                            update_log(f"Line {line_number}: Received: {stripped_line}")
                            log_data(RECEIVE_LOG_FILE, f"Line {line_number}: Received: {stripped_line}")

                        # Count occurrences of '%'
                        if stripped_line == "%":
                            percent_count += 1
                            if percent_count == 2:  # End of transmission after second '%'
                                update_log("Second '%' detected. Closing port and saving file.")
                                ser.close()
                                ACTIVE_PROCESS = None
                                update_gui_buttons()
                                return

                time.sleep(0.1)  # Check periodically

    except Exception as e:
        update_log(f"Error during file reception: {e}")
    finally:
        # Close the COM port at the end
        ser.close()
        update_log("COM port closed after receiving.")
        ACTIVE_PROCESS = None
        update_gui_buttons()

# GUI Application
def start_send():
    global FILENAME
    if ACTIVE_PROCESS is not None:
        update_log("Another process is active. Cannot start Send.")
        return

    file_path = filedialog.askopenfilename(filetypes=[["NC Files", "*.nc"]])
    if not file_path:
        return

    FILENAME = file_path
    update_log(f"Selected file: {FILENAME}")

    send_thread = threading.Thread(target=send_file)
    send_thread.start()
    update_gui_buttons()

def preview_lines(file_lines, index, preview_range=5):
    """
    Generate a preview of lines surrounding the specified index.

    :param file_lines: List of lines from the file.
    :param index: The line index to preview around (0-based).
    :param preview_range: Number of lines to show before and after.
    :return: Preview text.
    """
    start = max(0, index - preview_range)
    end = min(len(file_lines), index + preview_range + 1)
    return "\n".join(f"{i + 1}: {line.strip()}" for i, line in enumerate(file_lines[start:end], start=start))

def custom_input_dialog(file_lines, max_line, max_index=None):
    """
    Custom dialog to input starting line or Nxxx index with separate fields, dynamic previews,
    and enhanced handling for intermediate Nxxx indices.

    :param file_lines: List of lines from the file.
    :param max_line: Maximum line number.
    :param max_index: Maximum Nxxx index (if applicable).
    :return: (start_line, start_index) tuple.
    """
    dialog = tk.Toplevel(root)
    dialog.title("Partial Send - Specify Starting Point")
    dialog.geometry("600x500")
    dialog.resizable(False, False)

    selected_line = tk.IntVar(value=2)  # Default to line 2
    selected_index = tk.IntVar(value=None)

    # Preprocess file to find all Nxxx values
    n_values = []
    if max_index:
        for i, line in enumerate(file_lines):
            stripped_line = line.lstrip()
            if stripped_line.startswith("N"):
                try:
                    n_value = int(stripped_line[1:].split()[0])  # Extract the number after N
                    n_values.append(n_value)
                except ValueError:
                    continue

    n_min = min(n_values) if n_values else None
    n_max = max(n_values) if n_values else None

    # Line Number Section
    tk.Label(dialog, text=f"Enter line number (2-{max_line}):").pack(pady=5, anchor=tk.W)
    line_entry = tk.Entry(dialog)
    line_entry.pack(fill=tk.X, padx=10, pady=5)

    # Nxxx Index Section (Optional)
    index_entry = None
    tooltip = None

    if max_index:
        tk.Label(dialog, text=f"Enter Nxxx index (min {n_min}, max {n_max}):").pack(pady=5, anchor=tk.W)
        index_entry = tk.Entry(dialog)
        index_entry.pack(fill=tk.X, padx=10, pady=5)

        # Create tooltip for invalid input
        tooltip = tk.Label(dialog, text="", fg="red", font=("Arial", 10, "italic"))
        tooltip.pack(pady=5)

    # Preview Window
    preview_text = tk.Text(dialog, height=15, wrap=tk.WORD, state=tk.DISABLED)
    preview_text.pack(fill=tk.BOTH, padx=10, pady=10, expand=True)

    # OK Button
    ok_button = tk.Button(dialog, text="Confirm", state=tk.DISABLED)
    ok_button.pack(pady=10)

    # Cancel Button
    tk.Button(dialog, text="Cancel", command=dialog.destroy).pack(pady=5)

    def update_line_preview(*args):
        """
        Update preview based on the entered line number and disable the Nxxx field.
        """
        if line_entry.get():
            index_entry.config(state=tk.DISABLED)
        else:
            index_entry.config(state=tk.NORMAL)

        try:
            line_valid = False
            if line_entry.get():
                line_num = int(line_entry.get())
                if 2 <= line_num <= max_line:
                    line_valid = True
                    preview = preview_lines(file_lines, line_num - 1)
                    preview_text.config(state=tk.NORMAL)
                    preview_text.delete(1.0, tk.END)
                    preview_text.insert(tk.END, preview)
                    preview_text.config(state=tk.DISABLED)
            if not line_valid:
                raise ValueError
        except ValueError:
            preview_text.config(state=tk.NORMAL)
            preview_text.delete(1.0, tk.END)
            preview_text.insert(tk.END, "Invalid line number. Enter a value between 2 and the maximum.")
            preview_text.config(state=tk.DISABLED)

        validate_inputs()

    def update_index_preview(*args):
        """
        Update preview based on the entered Nxxx index, disable the line number field,
        and display nearest valid value for invalid input.
        """
        if index_entry.get():
            line_entry.config(state=tk.DISABLED)
        else:
            line_entry.config(state=tk.NORMAL)

        try:
            index_valid = False
            tooltip_text = ""
            if max_index and index_entry.get():
                index_num = int(index_entry.get())
                if n_min <= index_num <= n_max:
                    if index_num in n_values:  # Exact match
                        index_valid = True
                        for i, line in enumerate(file_lines):
                            if line.lstrip().startswith(f'N{index_num}'):
                                preview = preview_lines(file_lines, i)
                                preview_text.config(state=tk.NORMAL)
                                preview_text.delete(1.0, tk.END)
                                preview_text.insert(tk.END, preview)
                                preview_text.config(state=tk.DISABLED)
                                break
                    else:  # Approximate match: Use nearest valid value for the preview
                        closest_lower = max([n for n in n_values if n < index_num], default=None)
                        closest_higher = min([n for n in n_values if n > index_num], default=None)
                        nearest = closest_lower if closest_lower else closest_higher
                        tooltip_text = (
                            f"N{index_num} does not exist.\n"
                            f"Nearest valid index: N{nearest}"
                        )
                        # Display preview for the nearest valid index
                        for i, line in enumerate(file_lines):
                            if line.lstrip().startswith(f'N{nearest}'):
                                preview = preview_lines(file_lines, i)
                                preview_text.config(state=tk.NORMAL)
                                preview_text.delete(1.0, tk.END)
                                preview_text.insert(tk.END, preview)
                                preview_text.config(state=tk.DISABLED)
                                break
                if not index_valid:
                    raise ValueError
        except ValueError:
            tooltip.config(text=tooltip_text, fg="red")
            index_entry.config(bg="lightcoral")
        else:
            tooltip.config(text="")
            index_entry.config(bg="white")

        validate_inputs()

    def validate_inputs():
        """
        Enable or disable the OK button based on valid input in either field.
        """
        line_valid = line_entry.get().isdigit() and 2 <= int(line_entry.get()) <= max_line
        index_valid = (
            max_index
            and index_entry.get()
            and index_entry.get().isdigit()
            and int(index_entry.get()) in n_values
        )
        if line_valid or index_valid:
            ok_button.config(state=tk.NORMAL)
        else:
            ok_button.config(state=tk.DISABLED)

    def confirm():
        """
        Confirm the input and close the dialog.
        """
        try:
            line_value = int(line_entry.get()) if line_entry.get() else None
            index_value = int(index_entry.get()) if max_index and index_entry.get() else None
            selected_line.set(line_value)
            selected_index.set(index_value)
            dialog.destroy()
        except ValueError:
            tk.messagebox.showerror("Error", "Please provide valid input.")

    # Bind validation and preview updates
    line_entry.bind("<KeyRelease>", update_line_preview)
    if max_index:
        index_entry.bind("<KeyRelease>", update_index_preview)
    ok_button.config(command=confirm)

    dialog.wait_window()

    return selected_line.get(), selected_index.get()


def start_partial_send():
    global FILENAME
    if ACTIVE_PROCESS is not None:
        update_log("Another process is active. Cannot start Partial Send.")
        return

    # Ask user to select a file
    file_path = filedialog.askopenfilename(filetypes=[("NC Files", "*.nc")])
    if not file_path:
        update_log("No file selected for partial send.")
        return

    FILENAME = file_path
    update_log(f"Selected file for partial send: {FILENAME}")

    # Load the file to calculate constraints
    try:
        with open(FILENAME, 'r') as file:
            lines = file.readlines()
            total_lines = len(lines)
            update_log(f"Loaded file: {FILENAME} with {total_lines} lines.")

            # Find the maximum valid line (before the second %)
            max_idx = total_lines - 1
            for i, line in enumerate(lines):
                if line.strip() == "%" and i > 0:  # Second % found
                    max_idx = i
                    break

            max_line = max_idx + 1
            max_index = None

            # Find the maximum Nxxx index if applicable
            for i, line in enumerate(lines[:max_idx + 1]):  # Only consider lines before the second %
                stripped_line = line.lstrip()
                if stripped_line.startswith("N"):
                    try:
                        index_value = int(stripped_line[1:].split()[0])  # Extract the number after N
                        if max_index is None or index_value > max_index:
                            max_index = index_value
                    except ValueError:
                        continue

            update_log(f"Maximum valid line: {max_line}, Maximum valid Nxxx: {max_index if max_index else 'None'}")

    except Exception as e:
        update_log(f"Error loading file: {e}")
        return

    # Show the custom input dialog
    start_line, start_index = custom_input_dialog(lines, max_line, max_index)

    # If neither start_line nor start_index is provided, cancel
    if not start_line and not start_index:
        update_log("No starting point provided. Cancelling partial send.")
        return

    # Start the partial send in a new thread
    partial_send_thread = threading.Thread(target=partial_send_file, args=(start_line, start_index))
    partial_send_thread.start()
    update_gui_buttons()


def start_receive():
    if ACTIVE_PROCESS is not None:
        update_log("Another process is active. Cannot start Receive.")
        return

    receive_thread = threading.Thread(target=receive_file)
    receive_thread.start()
    update_gui_buttons()

def stop_operations():
    global STOP_REQUESTED
    STOP_REQUESTED = True
    update_log("Stop button pressed.")
    update_gui_buttons()

def toggle_cycle():
    global CYCLE_SEND
    CYCLE_SEND = not CYCLE_SEND
    update_log(f"Cycle Send set to: {CYCLE_SEND}")

from serial.tools import list_ports

def open_settings():
    def save_and_close():
        try:
            settings["COM_PORT"]        = com_port_var.get()
            settings["BAUD_RATE"]       = int(baud_rate_entry.get())
            settings["DATA_BITS"]       = int(data_bits_entry.get())
            settings["STOP_BITS"]       = int(stop_bits_entry.get())
            settings["PARITY"]          = parity_var.get()
            settings["FLOW_CONTROL"]    = flow_control_var.get()
            settings["LOGGING_ENABLED"] = logging_var.get()       # True/False
            settings["TRANSMISSION"]    = transmission_var.get()  # True/False

            save_settings()

            # Update global variables
            global COM_PORT, BAUD_RATE, DATA_BITS, STOP_BITS, PARITY, LOGGING_ENABLED, TRANSMISSION
            COM_PORT        = settings["COM_PORT"]
            BAUD_RATE       = settings["BAUD_RATE"]
            DATA_BITS       = settings["DATA_BITS"]
            STOP_BITS       = settings["STOP_BITS"]
            PARITY          = settings["PARITY"]
            LOGGING_ENABLED = settings["LOGGING_ENABLED"]
            TRANSMISSION    = settings["TRANSMISSION"]

            update_log("Settings saved successfully.")
            settings_window.destroy()
        except ValueError:
            update_log("Error: Invalid input in settings fields.")

    # Settings window UI elements
    settings_window = tk.Toplevel(root)
    settings_window.title("Settings")
    settings_window.geometry("400x350")
    settings_window.resizable(False, False)

    # COM Port Dropdown
    tk.Label(settings_window, text="COM Port:").grid(row=0, column=0, sticky=tk.W, padx=10, pady=5)
    com_ports = [port.device for port in list_ports.comports()]
    if not com_ports:
        com_ports = ["No COM ports available"]
    com_port_var = tk.StringVar(value=settings["COM_PORT"])
    com_port_menu = tk.OptionMenu(settings_window, com_port_var, *com_ports)
    com_port_menu.grid(row=0, column=1, padx=10, pady=5)

    tk.Label(settings_window, text="Baud Rate:").grid(row=1, column=0, sticky=tk.W, padx=10, pady=5)
    baud_rate_entry = tk.Entry(settings_window)
    baud_rate_entry.insert(0, str(settings["BAUD_RATE"]))
    baud_rate_entry.grid(row=1, column=1, padx=10, pady=5)

    tk.Label(settings_window, text="Data Bits:").grid(row=2, column=0, sticky=tk.W, padx=10, pady=5)
    data_bits_entry = tk.Entry(settings_window)
    data_bits_entry.insert(0, str(settings["DATA_BITS"]))
    data_bits_entry.grid(row=2, column=1, padx=10, pady=5)

    tk.Label(settings_window, text="Stop Bits:").grid(row=3, column=0, sticky=tk.W, padx=10, pady=5)
    stop_bits_entry = tk.Entry(settings_window)
    stop_bits_entry.insert(0, str(settings["STOP_BITS"]))
    stop_bits_entry.grid(row=3, column=1, padx=10, pady=5)

    tk.Label(settings_window, text="Parity:").grid(row=4, column=0, sticky=tk.W, padx=10, pady=5)
    parity_var = tk.StringVar(value=settings["PARITY"])
    parity_menu = tk.OptionMenu(settings_window, parity_var, "NONE", "EVEN", "ODD", "MARK", "SPACE")
    parity_menu.grid(row=4, column=1, padx=10, pady=5)

    # Flow Control Dropdown
    tk.Label(settings_window, text="Flow Control:").grid(row=5, column=0, sticky=tk.W, padx=10, pady=5)
    flow_control_var = tk.StringVar(value=settings.get("FLOW_CONTROL", "Software"))
    flow_control_menu = tk.OptionMenu(settings_window, flow_control_var, "Software", "Hardware")
    flow_control_menu.grid(row=5, column=1, padx=10, pady=5)

    logging_var = tk.BooleanVar(value=settings["LOGGING_ENABLED"])
    logging_checkbox = tk.Checkbutton(settings_window, text="Enable Logging", variable=logging_var)
    logging_checkbox.grid(row=6, column=0, columnspan=2, padx=10, pady=10)

    transmission_var = tk.BooleanVar(value=settings["TRANSMISSION"])
    transmission_checkbox = tk.Checkbutton(settings_window, text="Wait for transmission start", variable=transmission_var)
    transmission_checkbox.grid(row=7, column=0, columnspan=2, padx=10, pady=10)

    # Save and Close Button
    save_button = tk.Button(settings_window, text="Save and Close", command=save_and_close)
    save_button.grid(row=8, column=0, columnspan=2, pady=20)
    
def open_eula():
    """
    Opens a scrollable Toplevel window displaying the End-User License Agreement.
    Adjust the text as needed for your specific requirements.
    """
    eula_window = tk.Toplevel(root)
    eula_window.title("End-User License Agreement")
    eula_window.geometry("500x400")
    eula_window.resizable(False, False)

    frame_eula = tk.Frame(eula_window)
    frame_eula.pack(fill=tk.BOTH, expand=True)

    eula_text = tk.Text(frame_eula, wrap=tk.WORD)
    eula_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

    scrollbar = tk.Scrollbar(frame_eula, command=eula_text.yview)
    scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
    eula_text.config(yscrollcommand=scrollbar.set)

    # --- Generic EULA for free personal use, commercial use by request ---
    eula_content = """\
END-USER LICENSE AGREEMENT (EULA)

Last Updated: [Date]

This End-User License Agreement (“Agreement”) is a legal agreement between you (“Licensee”) 
and [Developer Name] (“Developer”) for the use of the software application known as 
“[Software Name]” (“Software”).

By installing, downloading, copying, or otherwise using the Software, you agree to be 
bound by the terms of this Agreement. If you do not agree, do not install or use the Software.

1. License Grant
   Developer grants you a non-exclusive, non-transferable, revocable license to install and 
   use the Software solely for personal, non-commercial purposes, free of charge.

2. Commercial Use
   Any use of the Software for commercial, professional, revenue-generating, or for-profit 
   activities requires a separate commercial license from Developer. To obtain such a license, 
   you must contact Developer via the contact information provided on the official website or 
   through official communication channels.

3. Restrictions
   You shall not:
   - Reverse engineer, decompile, or disassemble the Software, except to the extent 
     permitted by applicable law.
   - Reproduce, distribute, sell, lease, or otherwise make the Software available to 
     any third party without explicit written consent from Developer.
   - Modify or create derivative works based on the Software, except where expressly 
     allowed by this Agreement or applicable law.

4. Ownership and Intellectual Property
   All rights, title, and interest in and to the Software, including but not limited to 
   any trademarks, copyrights, or other intellectual property rights, remain with Developer.

5. Disclaimer of Warranties
   THE SOFTWARE IS PROVIDED “AS IS,” WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING 
   BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE, AND 
   NON-INFRINGEMENT.

6. Limitation of Liability
   TO THE MAXIMUM EXTENT PERMITTED BY LAW, IN NO EVENT SHALL DEVELOPER BE LIABLE FOR ANY DIRECT, 
   INDIRECT, SPECIAL, INCIDENTAL, OR CONSEQUENTIAL DAMAGES ARISING OUT OF THE USE OR INABILITY 
   TO USE THE SOFTWARE.

7. Termination
   This Agreement is effective upon your acceptance or use of the Software and remains in effect 
   until terminated. Developer may terminate this Agreement at any time if you breach any provision. 
   Upon termination, you must immediately cease all use of the Software and destroy any copies.

8. Governing Law
   This Agreement shall be governed by and construed in accordance with the laws of [Applicable Jurisdiction].

9. Contact Information
   For questions about this Agreement, or to request a commercial license, please visit our website 
   or contact us at:
   - Website: http://u-solutions.eu
   - Practical Machinist Profile: https://www.practicalmachinist.com/forum/members/usolutions.242439/

BY INSTALLING OR USING THIS SOFTWARE, YOU ACKNOWLEDGE THAT YOU HAVE READ AND UNDERSTOOD 
THIS AGREEMENT AND AGREE TO BE BOUND BY ITS TERMS.
"""

    eula_text.insert(tk.END, eula_content)

    close_button = tk.Button(eula_window, text="Accept", command=eula_window.destroy)
    close_button.pack(pady=10)

# GUI Initialization
if getattr(sys, 'frozen', False):  # Running as a PyInstaller executable
    base_path = sys._MEIPASS
else:
    base_path = os.path.dirname(__file__)

icon_path = os.path.join(base_path, "uDNC.ico")

root = tk.Tk()
root.title("U-Solutions uDNC")
root.geometry("600x400")
try:
    root.iconbitmap(icon_path)
except Exception as e:
    print(f"Could not set icon: {e}")

# Frames and layout
frame_controls = tk.Frame(root)
frame_controls.pack(pady=10)

frame_log = tk.Frame(root)
frame_log.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

# Buttons
btn_send = tk.Button(frame_controls, text="Send File", command=start_send)
btn_send.grid(row=0, column=0, padx=10)

btn_receive = tk.Button(frame_controls, text="Receive File", command=start_receive)
btn_receive.grid(row=0, column=2, padx=10)

btn_partial_send = tk.Button(frame_controls, text="Partial Send", command=start_partial_send)
btn_partial_send.grid(row=0, column=1, padx=10)

btn_stop = tk.Button(frame_controls, text="Stop", command=stop_operations, state=tk.DISABLED)
btn_stop.grid(row=0, column=3, padx=10)

btn_settings = tk.Button(frame_controls, text="Settings", command=open_settings)
btn_settings.grid(row=0, column=4, padx=10)

cycle_send_var = tk.BooleanVar()
# Create "Cycle Send" checkbox
checkbox_cycle = tk.Checkbutton(frame_controls, text="Cycle Send", command=toggle_cycle)
checkbox_cycle.grid(row=1, column=0, padx=10)

# Create a Toplevel tooltip (for "Cycle Send" hover)
tooltip_cycle_send = tk.Toplevel(root)
tooltip_cycle_send.withdraw()  # Hide the tooltip initially
tooltip_cycle_send.overrideredirect(True)  # Remove window decorations
tooltip_label = tk.Label(
    tooltip_cycle_send,
    text="If enabled, the same file will be sent again after completion, awaiting XON signal.",
    bg="yellow",
    wraplength=300,
    relief=tk.SOLID,
    bd=1,
    padx=5,
    pady=5,
)
tooltip_label.pack()

def show_tooltip(event):
    # Position the tooltip near the checkbox
    tooltip_cycle_send.geometry(f"+{event.x_root + 10}+{event.y_root + 10}")
    tooltip_cycle_send.deiconify()  # Show the tooltip

def hide_tooltip(event):
    tooltip_cycle_send.withdraw()  # Hide the tooltip

# Bind tooltip to "Cycle Send" checkbox
checkbox_cycle.bind("<Enter>", show_tooltip)
checkbox_cycle.bind("<Leave>", hide_tooltip)

# Log Text Widget
log_text = tk.Text(frame_log, wrap=tk.WORD, height=15)
log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

scrollbar = tk.Scrollbar(frame_log, command=log_text.yview)
scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
log_text.config(yscrollcommand=scrollbar.set)

def open_link(url):
    webbrowser.open_new_tab(url)

# Add hyperlinks below the scrolling log
frame_links = tk.Frame(root)
frame_links.pack(pady=5)

link1 = tk.Label(frame_links, text="Contact us", fg="blue", cursor="hand2")
link1.pack(side=tk.LEFT, padx=10)
link1.bind("<Button-1>", lambda e: open_link("http://u-solutions.eu"))

link2 = tk.Label(frame_links, text="Practical Machinist forum profile", fg="blue", cursor="hand2")
link2.pack(side=tk.LEFT, padx=10)
link2.bind("<Button-1>", lambda e: open_link("https://www.practicalmachinist.com/forum/members/usolutions.242439/"))

# EULA link
link3 = tk.Label(frame_links, text="EULA", fg="blue", cursor="hand2")
link3.pack(side=tk.LEFT, padx=10)
link3.bind("<Button-1>", lambda e: open_eula())

# Update GUI Buttons
def update_gui_buttons():
    if ACTIVE_PROCESS == "send":
        btn_send.config(state=tk.DISABLED)
        btn_receive.config(state=tk.DISABLED)
        btn_stop.config(state=tk.NORMAL)
    elif ACTIVE_PROCESS == "receive":
        btn_send.config(state=tk.DISABLED)
        btn_receive.config(state=tk.DISABLED)
        btn_stop.config(state=tk.NORMAL)
    else:
        btn_send.config(state=tk.NORMAL)
        btn_receive.config(state=tk.NORMAL)
        btn_stop.config(state=tk.DISABLED)

# Run the GUI
update_gui_buttons()
root.mainloop()
