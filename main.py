import csv
import openpyxl
import pandas as pd


"""
Required fields:
Item Name
Item Group
Item Category
Default Price
POS Orders Print At

Fields for folders:
Item Folder (1 for true, 0 for false)
Belongs To Item Folder 
"""


def create_folder(item, options, suboptions):
    with open('testings.csv', 'a', newline='') as csvfile:
        fieldnames = ['Menu Item Full Name', 'Menu Item Group', 'Menu Item Category', 'Default Price', 'Dine In Price', 'Bar Price', 'Pick Up Price', 'Take Out Price', 'Delivery Price', 'Open Price Item', 'POS Orders Print At', 'Tax 1', 'Tax 2', 'Tax 3', 'This Is A Bar Item', 'This Is A Weighted Item', 'Tare', 'Barcode', 'Item Folder', 'Belongs To Item Folder']
        writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
        writer.writeheader()
    
            # For the folder menu items
        for option in options:
            folder_name = f"{option} {item}"
            print(f"{folder_name}")

                # For the subfolders menu items
            for sub in suboptions:
                subfolder_name = f"{folder_name} - {sub}"
                print(f"{subfolder_name}")
                writer.writerow({'Menu Item Full Name': subfolder_name, 'Menu Item Group': 'Coffee', 'Menu Item Category': 'Drinks', 'Default Price': '1', 'POS Orders Print At': 'N', 'Item Folder': '1', 'Belongs To Item Folder': folder_name})
                


def read_csv():
    with open('testings.csv', mode='r') as csvfile:
        return csv.DictReader(csvfile)
        for row in reader:
            print(row)
            #print(f"{row['Menu Item Full Name']} | {row['Menu Item Group']} | {row['Menu Item Category']} | $ {row['Default Price']} ")

def conv_pd():
    df = pd.read_csv('testings.csv')
    print(df)
    return df


# Main name / Type name / Size
#create_folder("Coffee", ["Hot", "Cold"], ["12oz", "16oz", "20oz"])
#create_folder("Tea", ["Hot", "Cold"], ["12oz", "16oz", "20oz"])
conv_pd()
#conv_csv_to_excel('testings.csv', 'testings.xlsx')
