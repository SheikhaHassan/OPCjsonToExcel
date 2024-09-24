import json
import openpyxl
from openpyxl import Workbook

#open and read JSON file
with open('OPC.json') as jsonFile:
    json_data = json.load(jsonFile)
(len(channelnames))


def write_to_excel(data, filename):
    # Create a new workbook
    wb = openpyxl.Workbook()
    ws = wb.active

    # Write data to the worksheet
    for row_idx, row_data in enumerate(data, start=1):
        for col_idx, cell_value in enumerate(row_data, start=1):
            ws.cell(row=row_idx, column=col_idx, value=cell_value)

    # Save the workbook
    wb.save(filename)



def extract_ip_addresses(devices):
    ip_addresses = []
    for device in devices:
        if "servermain.DEVICE_ETHERNET_COMMUNICATIONS_IP" in device:
            ip_addresses.append(device["servermain.DEVICE_ETHERNET_COMMUNICATIONS_IP"])
    return ip_addresses

def write_to_excel(json_data):
    # Create a new Workbook
    wb = Workbook()


    # Iterate over the range
    for i in range(453):
        # Create a new sheet for each iteration
        sheet = wb.create_sheet(title=f"Channel_{i + 1}")

        # Extract channel_name, device names, driver names, and IP addresses

        channel_name = json_data["project"]["channels"][i]["common.ALLTYPES_NAME"]
        devices = json_data["project"]["channels"][i]["devices"]
        device_names = [device["common.ALLTYPES_NAME"] for device in devices]
        driver_names = [device["servermain.MULTIPLE_TYPES_DEVICE_DRIVER"] for device in devices]
        ip_addresses = extract_ip_addresses(devices)

        # Write column names
        sheet['A1'] = 'channel Name'
        sheet['B1'] = 'Device Name'
        sheet['C1'] = 'Driver Name'
        sheet['D1'] = 'IP Address'

        # Write device names, driver names, and IP addresses into the sheet
        for row, (device_name, driver_name, ip_address) in enumerate(zip(device_names, driver_names, ip_addresses), start=2):
            sheet[f'A{row}'] = channel_name
            sheet[f'B{row}'] = device_name
            sheet[f'C{row}'] = driver_name
            sheet[f'D{row}'] = ip_address

    # Save the workbook
    wb.save("final.xlsx")

# Call the function on JSON data
write_to_excel(json_data)













