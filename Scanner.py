import os
import threading
#import pythoncom
import tkinter as tk
from tkinter import ttk,messagebox
from tkinter import *
import comtypes.client
import comtypes
import pytesseract
import PIL.Image
import re
import json
import shutil
import datetime
import requests
from pyshortcuts import make_shortcut

src_folder = os.getcwd()
dest_folder = "C:/Bill Extracted"

if not os.path.exists(dest_folder):
    os.makedirs(dest_folder)

def scan_document():
    comtypes.CoInitialize()
    try:
        wia = comtypes.client.CreateObject('WIA.CommonDialog')
        device = wia.ShowSelectDevice(1, False)

        if not device:
            print("No WIA device is connected!")
            return

        print(f"Connected device: {device.Properties('Name').Value}")
        img = device.Items[1].Transfer('{B96B3CAE-0728-11D3-9D7B-0000F81EF32E}')
        img.SaveFile('scanned_image.jpg')
        print("Document scanned successfully!")

    except comtypes.COMError as e:
        print("Error: No WIA device available. Please connect a scanner.")
        scanner_error1 = messagebox.showinfo("Error1","No Scanner is connected.")

    except Exception as e:
        print(f"An unexpected error occurred: {e}")
        scanner_error2 = messagebox.showinfo("Error2",f"Unexpected error : {e}")

    finally:
        comtypes.CoUninitialize()

def extract_key_value_pairs(text):
    pairs = {
        'vehicle_no': 'Not Entered',
        'material': 'Not Entered',
        'vendor_name': 'Not Entered',
        'gross_weight': 'Not Entered',
        'tare_weight': 'Not Entered',
        'net_weight': 'Not Entered',
        'in_date': 'Not Entered',
        'in_time': 'Not Entered',
        'status': 1,
        'user_id': 1,
        'apikey': '6XR2GMexyDYrvmwZnYNYy'
    }

    lines = text.splitlines()

    date_pattern = r"\b(\d{1,2})[/-](\d{1,2})[/-](\d{4})\b"
    vehicle_pattern = r"TN\s*([^\s]+)"
    name_pattern = r"(?i)\b(party|customer|client|name|m)\b\s*(.*)"
    time_pattern = r"([^\s]+)\s*[AP]M"
    kg_pattern = r"(\d+[,.]?\d*)\s*-?\s*Kg"
    weight_pattern=r"(?:weight|wt)\s*[:\-]?\s*(\d+(\.\d+)?)"
    stone_pattern = r"(.*?)[\s:]*(stone|ston|sto)\b"
    material_pattern = r"Material\s*[:\-]?\s*(.+)"

    weights = []

    for line in lines:
        segments = re.split(r'\s{2,}|\s(?=[A-Za-z]+\s*[:-])', line)

        for segment in segments:

            vehicle_match = re.search(vehicle_pattern, segment)
            if vehicle_match:
                pairs['vehicle_no'] = vehicle_match.group(0).strip()

            stone_match = re.search(stone_pattern, segment, re.IGNORECASE)
            if stone_match and 'material' not in pairs:
                material = stone_match.group(1).strip()
                pairs['material'] = f"{material} Stone"

            material_match = re.search(material_pattern, segment)
            if material_match:
                pairs['material'] = material_match.group(1).strip()

            name_match = re.search(name_pattern, segment)
            if name_match:
                pairs['vendor_name'] = name_match.group(2).strip()

            kg_matches = re.findall(kg_pattern, line)
            for match in kg_matches:
                if match:
                    weights.append(match.strip())

            if len(weights) >= 3:
                pairs['gross_weight'] = f"{weights[0]} Kg"
                pairs['tare_weight'] = f"{weights[1]} Kg"
                pairs['net_weight'] = f"{weights[2]} Kg"

            else:
                weight_match = re.search(weight_pattern, line, re.IGNORECASE)

                if weight_match and all(key not in pairs for key in ['gross_weight', 'tare_weight', 'net_weight']):
                    if weight_match.group(1):
                        weights.append(weight_match.group(1).strip())

                    if len(weights) >= 3:
                        pairs['gross_weight'] = f"{weights[0]} Kg"
                        pairs['tare_weight'] = f"{weights[1]} Kg"
                        pairs['net_weight'] = f"{weights[2]} Kg"

            date_match = re.search(date_pattern, segment)
            if date_match:
                pairs['in_date'] = date_match.group(0).strip()

            time_match = re.search(time_pattern, line)
            if time_match:
                pairs['in_time'] = time_match.group(0)

    return pairs

scan_complete_event = threading.Event()

def execute_program():
    try:
        scan_document()

        global api_url, api_key, myconfig,text,key_value_pairs,json_data

        api_url = "https://crusher.rndhub.in/api/postcrusherstone"
        api_key = "6XR2GMexyDYrvmwZnYNYy"

        myconfig = r"--psm 6 --oem 3"
        pytesseract.pytesseract.tesseract_cmd = r"C:\tesseract\tesseract.exe"
        text = pytesseract.image_to_string(PIL.Image.open("scanned_image.jpg"), config=myconfig)
        print(text)

        key_value_pairs = extract_key_value_pairs(text)
        json_data = json.dumps(key_value_pairs, indent=4)
        print(json_data)

        for filename in os.listdir(src_folder):
            file_path = os.path.join(src_folder, filename)

            if os.path.isfile(file_path) and filename.lower().endswith(('.png', '.jpg', '.jpeg', '.gif')):
                current_time = datetime.datetime.now().strftime("%Y-%m-%d %H-%M-%S")
                new_filename = f"{current_time}.png"
                shutil.move(file_path, os.path.join(dest_folder, new_filename))

        scan_complete_event.set()

    except Exception as e:
        print(f"An unwanted error occurred: {e}")
        scanner_error3=messagebox.showinfo("Error3",f"An unwanted error : {e}")

def on_scan_complete():
    #progress_bar.stop()
    #progress_bar.pack_forget()
    scanning_text.destroy()
    json_data_parsed = json.loads(json_data)

    vendor_name = json_data_parsed.get('vendor_name', 'Not Entered')
    vendor_json = Label(root, text="Vendor Name :", font=('arial', 14, 'normal'), fg="white", background="black")
    vendor_json.place(x=10, y=10)
    vendor_input = tk.Entry(root, font=('arial', 14, 'normal'))
    vendor_input.insert(0, vendor_name)
    vendor_input.place(x=200, y=10)

    vehicle_no = json_data_parsed.get('vehicle_no', 'Not Entered')
    vehicle_json = Label(root, text="Vehicle Number :", font=('arial', 14, 'normal'), fg="white", background="black")
    vehicle_json.place(x=10, y=60)
    vehicle_input = tk.Entry(root, font=('arial', 14, 'normal'))
    vehicle_input.insert(0, vehicle_no)
    vehicle_input.place(x=200, y=60)

    material = json_data_parsed.get('material', 'Not Entered')
    material_json = Label(root, text="Material :", font=('arial', 14, 'normal'), fg="white", background="black")
    material_json.place(x=10, y=110)
    material_input = tk.Entry(root, font=('arial', 14, 'normal'))
    material_input.insert(0, material)
    material_input.place(x=200, y=110)

    gross_weight = json_data_parsed.get('gross_weight', 'Not Entered')
    gross_weight_json = Label(root, text="Gross Weight :", font=('arial', 14, 'normal'), fg="white", background="black")
    gross_weight_json.place(x=10, y=160)
    gross_weight_input = tk.Entry(root, font=('arial', 14, 'normal'))
    gross_weight_input.insert(0, gross_weight)
    gross_weight_input.place(x=200, y=160)

    tare_weight = json_data_parsed.get('tare_weight', 'Not Entered')
    tare_weight_json = Label(root, text="Tare Weight :", font=('arial', 14, 'normal'), fg="white", background="black")
    tare_weight_json.place(x=10, y=210)
    tare_weight_input = tk.Entry(root, font=('arial', 14, 'normal'))
    tare_weight_input.insert(0, tare_weight)
    tare_weight_input.place(x=200, y=210)

    net_weight = json_data_parsed.get('net_weight', 'Not Entered')
    net_weight_json = Label(root, text="Net Weight :", font=('arial', 14, 'normal'), fg="white", background="black")
    net_weight_json.place(x=10, y=260)
    net_weight_input = tk.Entry(root, font=('arial', 14, 'normal'))
    net_weight_input.insert(0, net_weight)
    net_weight_input.place(x=200, y=260)

    in_date = json_data_parsed.get('in_date', 'Not Entered')
    in_date_json = Label(root, text="In Date :", font=('arial', 14, 'normal'), fg="white", background="black")
    in_date_json.place(x=10, y=310)
    in_date_input = tk.Entry(root, font=('arial', 14, 'normal'))
    in_date_input.insert(0, in_date)
    in_date_input.place(x=200, y=310)

    in_time = json_data_parsed.get('in_time', 'Not Entered')
    in_time_json = Label(root, text="In Time :", font=('arial', 14, 'normal'), fg="white", background="black")
    in_time_json.place(x=10, y=360)
    in_time_input = tk.Entry(root, font=('arial', 14, 'normal'))
    in_time_input.insert(0, in_time)
    in_time_input.place(x=200, y=360)

    def update_details():
        json_data_parsed['vendor_name'] = vendor_input.get()
        json_data_parsed['vehicle_no'] = vehicle_input.get()
        json_data_parsed['material'] = material_input.get()
        json_data_parsed['gross_weight'] = gross_weight_input.get()
        json_data_parsed['tare_weight'] = tare_weight_input.get()
        json_data_parsed['net_weight'] = net_weight_input.get()
        json_data_parsed['in_date'] = in_date_input.get()
        json_data_parsed['in_time'] = in_time_input.get()
        json_data_updated = json.dumps(json_data_parsed, indent=4)
        print("Updated JSON Data:\n", json_data_updated)

        headers = {
            "Authorization": f"Bearer {api_key}",
            "Content-Type": "application/json"
        }

        response = requests.post(api_url, headers=headers, data=json_data)

        if response.status_code == 200:
            print("Data successfully posted")
            print("Response:", response.json())
        else:
            print(f"Failed to post data: {response.status_code}")
            print("Response:", response.text)

        #move()

        root.after(0, lambda: on_scan_complete())

        user_choice = messagebox.askquestion("Successfully Scanned",
                                             "Bill scanned successfully! Do you want to scan the next one?",
                                             icon='question')

        if user_choice == 'yes':
            global scan_count
            scan_count=0
            run_program_in_thread()
        else:
            root.destroy()

    update_button = tk.Button(root, text="Process", command=update_details, font=('arial', 14, 'normal'))
    update_button.place(x=330, y=400)

scan_count=0

def run_program_in_thread():
    global scan_count, progress_bar, scanning_text,scanning_window

    scanning_window = tk.Toplevel(root)
    scanning_window.title("Scanning")
    scanning_window.geometry("70x450")
    scanning_window.configure(bg="black")
    scanning_window.attributes("-topmost", True)
    center_window(scanning_window, window_width, window_height)

    if scan_count == 0:
        progress_bar = ttk.Progressbar(scanning_window, orient="horizontal", mode="determinate", length=300)
        scanning_text = tk.Label(scanning_window, text="Scanning...", background="black", fg="white", font=12)
        scanning_text.place(y=230, x=225)
        progress_bar.pack(pady=200)
        progress_bar.start(235)

    scan_thread = threading.Thread(target=execute_program)
    scan_thread.start()

    check_scan_completion()


def check_scan_completion():
    if scan_complete_event.is_set():
        progress_bar.stop()
        progress_bar.pack_forget()
        scanning_text.destroy()

        scan_complete_event.clear()

        on_scan_complete()

        scanning_window.destroy()
    else:
        root.after(100, check_scan_completion)


def on_proceed_click():
    proceed_button.destroy()
    company_name.destroy()
    run_program_in_thread()

def create_desktop_shortcut():
    desktop_path = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop')

    exe_path = os.getcwd()

    make_shortcut(exe_path,
                  name='Crusher Scanner',
                  description='A Python Application',
                  desktop=True,
                  icon="C:/Users/saipa/Downloads/342170db2486b7b8e7a3d23b88aae788.ico")

def center_window(window, width, height):
    screen_width = window.winfo_screenwidth()
    screen_height = window.winfo_screenheight()
    x_coordinate = int((screen_width / 2) - (width / 2))
    y_coordinate = int((screen_height / 2) - (height / 2))
    window.geometry(f"{width}x{height}+{x_coordinate}+{y_coordinate}")

root = tk.Tk()
root.title("Crusher Scanner")
root.configure(bg="black")
window_width = 750
window_height = 450
center_window(root, window_width, window_height)

proceed_button = tk.Button(root,text="Click here to start scanning",activebackground="blue",activeforeground="white",anchor="center",bd=3,bg="lightgray",cursor="hand2",disabledforeground="gray",fg="black",font=("Arial", 12),height=2,highlightbackground="black",highlightcolor="green",highlightthickness=2,justify="center",overrelief="raised",padx=10,pady=5,width=15,wraplength=100,command=on_proceed_click)
proceed_button.pack(pady=180)

progress_bar = ttk.Progressbar(root, orient="horizontal", mode="determinate", length=300)

scanning_text = Label(root, text="Scanning...",background="black",fg="white",font=12)

company_name = Label(root, text="Welcome to Crushers Stone Scanning",font=16,fg="white",background="black")
company_name.place(x=200,y=50)

root.mainloop()