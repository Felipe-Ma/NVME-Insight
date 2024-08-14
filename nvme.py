import subprocess
import openpyxl
from openpyxl import Workbook
from datetime import datetime

# Function to get NVMe device list
def get_nvme_devices():
    result = subprocess.run(['nvme', 'list'], stdout=subprocess.PIPE, stderr=subprocess.PIPE)
    if result.returncode != 0:
        print(f"Error listing NVMe devices: {result.stderr.decode('utf-8')}")
        return []
    
    devices = []
    lines = result.stdout.decode('utf-8').splitlines()
    for line in lines:
        if line.startswith('/dev/nvme'):
            devices.append(line.split()[0])
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
    for line in lines:
        if "percentage used" in line.lower():
            health_info['Percentage Used'] = line.split(":")[-1].strip()
        elif "temperature" in line.lower() and "sensor" not in line.lower():
            health_info['Temperature'] = line.split(":")[-1].strip()
        elif "critical warning" in line.lower():
            health_info['Critical Warning'] = line.split(":")[-1].strip()
        elif "available spare" in line.lower() and "threshold" not in line.lower():
            health_info['Available Spare'] = line.split(":")[-1].strip()
        elif "data units written" in line.lower():
            health_info['Data Units Written'] = line.split(":")[-1].strip()
        elif "data units read" in line.lower():
            health_info['Data Units Read'] = line.split(":")[-1].strip()
    return health_info

# Function to write the health info to an Excel file
def write_to_excel(device, health_info):
    headers = ['Timestamp', 'Device', 'Percentage Used', 'Temperature', 'Critical Warning', 'Available Spare', 'Data Units Written', 'Data Units Read']

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
        row = [datetime.now().strftime("%Y-%m-%d %H:%M:%S"), device]
        row.extend([health_info.get(key, 'N/A') for key in headers[2:]])
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
        print("No NVMe devices found.")
        return

    for device in devices:
        health_info = get_nvme_health(device)
        if health_info:
            write_to_excel(device, health_info)

if __name__ == "__main__":
    main()

