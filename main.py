import subprocess
import openpyxl
from openpyxl import Workbook
from datetime import datetime
import re

# Function to get NVMe device list
def get_nvme_devices():
    result = subprocess.run(['nvme', 'list'], stdout=subprocess.PIPE, stderr=subprocess.PIPE)
    if result.returncode != 0:
        print(f"Error listing NVMe devices: {result.stderr.decode('utf-8')}")
        return []
    
    devices = []
    lines = result.stdout.decode('utf-8').splitlines()
    for line in lines:
        # print(line)
        if line.startswith('/dev/nvme'):
            parts = line.split()
            device = parts[0]
            model_name = parts[3] 
            if model_name.startswith("SOLIDIGM"):
                devices.append(device)
    return devices

# Function to get NVMe health information
def get_nvme_health(device):
    try:
        result = subprocess.run(['nvme', 'smart-log', device], stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        output = result.stdout.decode('utf-8')
        if result.returncode != 0:
            print(f"Error: {result.stderr.decode('utf-8')}")
            return None
        return parse_nvme_output(output)
    except Exception as e:
        print(f"Error retrieving NVMe health for {device}: {e}")
        return None

# Function to parse the nvme-cli output
def parse_nvme_output(output):
    health_info = {}
    lines = output.splitlines()
    # print(lines)
    for line in lines:
        if "percentage_used" in line.lower():
            #print("\n\n\n" + line.lower)
            health_info['Percentage Used'] = line.split(":")[-1].strip()
        #elif "temperature" in line.lower() and "sensor" not in line.lower():
            #health_info['Temperature'] = line.split(":")[-1].strip()
        elif "critical_warning" in line.lower():
            health_info['Critical Warning'] = line.split(":")[-1].strip()
        elif "available_spare" in line.lower() and "threshold" not in line.lower():
            health_info['Available Spare'] = line.split(":")[-1].strip()
        elif "media_errors" in line.lower() and "threshold" not in line.lower():
            health_info['Media Errors'] = line.split(":")[-1].strip()
        elif "data units written" in line.lower():
            health_info['Data Units Written'] = line.split(":")[-1].strip()
        elif "data units read" in line.lower():
            health_info['Data Units Read'] = line.split(":")[-1].strip()
        elif "host_read_commands" in line.lower():
            health_info['Host Read Commands'] = line.split(":")[-1].strip()
        elif "host_write_commands" in line.lower():
            health_info['Host Write Commands'] = line.split(":")[-1].strip()
        elif "power_cycles" in line.lower():
            health_info['Power Cycles'] = line.split(":")[-1].strip()
        elif "power_on_hours" in line.lower():
            health_info['Power On Hours'] = line.split(":")[-1].strip()
        elif "unsafe_shutdowns" in line.lower():
            health_info['Unsafe Shutdowns'] = line.split(":")[-1].strip()
            
    return health_info

# Function to get NVMe device serial number and model number
def get_nvme_device_info(device):
    try:
        result = subprocess.run(['nvme', 'id-ctrl', device], stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        output = result.stdout.decode('utf-8')
        if result.returncode != 0:
            print(f"Error: {result.stderr.decode('utf-8')}")
            return "Unknown", "Unknown"
        lines = output.splitlines()
        serial_number = "Unknown"
        model_number = "Unknown"
        for line in lines:
            #print(line + "+++++")
            if "sn" in line.lower(): 
                serial_number = line.split(":")[-1].strip()
            if re.match(r'^\s*mn\s*:', line.lower()):
                #print("+ " + line)
                mn_key, value = line.split(":", 1)
                #print(mn_key, value)
                model_number = value
            if re.match(r'^\s*fr\s*:', line.lower()):
                #print("+ " + line)
                fr_key, value = line.split(":", 1)
                #print(fr_key, value)
                firmware_version = value

        return serial_number, model_number, firmware_version
    except Exception as e:
        print(f"Error retrieving NVMe serial number and model number for {device}: {e}")
        return "Unknown", "Unknown"

# Function to write the health info to an Excel file
def write_to_excel(device, serial_number, model_number, firmware_version, health_info):
    headers = ['Timestamp', 'Device', 'Serial Number', 'Model Number', 'Firwmare Version','Percentage Used', 'Critical Warning', 'Available Spare', 'Media Errors', 'Data Units Written', 'Data Units Read', 'Host Read Commands', 'Host Write Commands', 'Power Cycles', 'Power On Hours', 'Unsafe Shutdowns']
    #print(health_info)
    try:
        # Load or create the Excel file
        try:
            workbook = openpyxl.load_workbook('nvme_health_report.xlsx')
            sheet = workbook.active
        except FileNotFoundError:
            workbook = Workbook()
            sheet = workbook.active
            # Write headers
            sheet.append(headers)

        # Append data
        row = [datetime.now().strftime("%Y-%m-%d %H:%M:%S"), device, serial_number, model_number, firmware_version]
        row.extend([health_info.get(key, 'N/A') for key in headers[5:]])
        sheet.append(row)

        # Save the file
        workbook.save('nvme_health_report.xlsx')
        print(f"Data written to nvme_health_report.xlsx for device {device}")

    except Exception as e:
        print(f"Error writing to Excel: {e}")

# Main function to check health and document it
def main():
    devices = get_nvme_devices()
    if not devices:
        print("No Solidigm NVMe devices found.")
        return

    for device in devices:
        serial_number, model_number, firmware_version = get_nvme_device_info(device)
        health_info = get_nvme_health(device)
        if health_info:
            write_to_excel(device, serial_number, model_number, firmware_version, health_info)

if __name__ == "__main__":
    main()
