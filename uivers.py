import nicegui as ng
from nicegui import app, ui, events
import csv
import pandas as pd
from openpyxl import load_workbook

# contains info regarding Menu Items
class MenuItem:
    def __init__(self, name_type=None, group=None, category=None, price=None, folder=None):
        self.holder = []
        self.count = 0

        if name_type is not None:
            self.add_item(name_type, group, category, price, folder)
            

    def add_item(self, name_type:str, group:str, category:str, price:int, folder_name:str, print_loc:str):
        item = {
            'Menu Item Full Name': name_type,
            'Menu Item Group': group,
            'Menu Item Category': category,
            'Default Price': price,
            'Belongs To Item Folder': folder_name,
            'POS Orders Print At': print_loc,
        }
        self.holder.append(item)
        self.count += 1
        #print(self.holder)

    def delete_item(self, item1):
        self.holder.remove(item1)
        #print(self.holder)

    def update_item(self, id, val):
        self.holder[id]['Belongs To Item Folder'] = val

    def get_rows(self):
        return self.holder
# ------------------------End of Class MenuItem------------------------------------------------


#------------------------Start of Functions------------------------------------------------

def add_row():
    #ui.label(f"Added: {fn.value + '-' + it.value} | {ip.value} | {fn.value} | {ig.value} | {ic.value}").update()
    if len(fn.value) > 0:
        table_l.add_item(fn.value + ' ' + it.value, ig.value, ic.value, ip.value, fn.value, print_loc='N')
    else:
        table_l.add_item(it.value, ig.value, ic.value, ip.value, fn.value, print_loc='N')
    

# include limit of 32 per unique group / POS default menu setup
def add_table():
    add_row()
    grid.update_rows(table_l.get_rows())


def add():
    add_row()
    row = table_l.holder[-1]
    with grid.props.suspend_updates():
        grid.options['rowData'].append(row)
    grid.run_grid_method('applyTransaction', {'add': [row]})
    grid.run_grid_method('ensureIndexVisible', len(grid.options['rowData']) - 1)
    show_items()


async def export():
    data = await show_grid()
    wb = load_workbook("output.xlsx")
    ws = wb["Sheet1"]

    headers = [cell.value for cell in ws[1]]
    for row in data:
        ws.append([row.get(h, None) for h in headers])
    wb.save("output.xlsx")



async def show_items():
    row = await grid.get_client_data()
    print(row)


async def show_grid():
    return await grid.get_client_data()

     
#------------------------End of Functions------------------------------------------------



# Global variables go here

table_l = MenuItem()    # MenuItem may not be needed
printer_options = ['1', '2', '3', '4', '5', '6', '7', '8', '9', '10', 'B', 'N']

columns = [
    {'field': 'Menu Item Full Name'},
    {'field': 'Menu Item Group'},
    {'field': 'Menu Item Category'},
    {'field': 'Default Price', 'cellEditor': 'agNumberCellEditor','cellEditorParams': {'min': 0, 'max': 10000, 'precision': 2}},
    {'field': 'Dine In Price', 'hide': True, 'cellEditor': 'agNumberCellEditor','cellEditorParams': {'min': 0, 'max': 10000, 'precision': 2}},
    {'field': 'Bar Price', 'hide': True, 'cellEditor': 'agNumberCellEditor','cellEditorParams': {'min': 0, 'max': 10000, 'precision': 2}},
    {'field': 'Pick Up Price', 'hide': True, 'cellEditor': 'agNumberCellEditor','cellEditorParams': {'min': 0, 'max': 10000, 'precision': 2}},
    {'field': 'Take Out Price', 'hide': True, 'cellEditor': 'agNumberCellEditor','cellEditorParams': {'min': 0, 'max': 10000, 'precision': 2}},
    {'field': 'Delivery Price', 'hide': True, 'cellEditor': 'agNumberCellEditor','cellEditorParams': {'min': 0, 'max': 10000, 'precision': 2}},
    {'field': 'Open Price Item', 'hide': True},
    {'field': 'POS Orders Print At', 'cellEditor': 'agSelectCellEditor', 'cellEditorParams':{'values': printer_options}},
    {'field': 'Tax 1', 'cellRenderer': 'agCheckboxCellRenderer', 'cellEditor': 'agCheckboxCellEditor',},
    {'field': 'Tax 2', 'hide': True, 'cellRenderer': 'agCheckboxCellRenderer', 'cellEditor': 'agCheckboxCellEditor',},
    {'field': 'Tax 3', 'hide': True, 'cellRenderer': 'agCheckboxCellRenderer', 'cellEditor': 'agCheckboxCellEditor',},
    {'field': 'This Is A Bar Item', 'hide': True},
    {'field': 'This Is A Weighted Item', 'hide': True},
    {'field': 'Tare', 'hide': True},
    {'field': 'Barcode', 'hide': True},
    {'field': 'Item Folder', 'cellRenderer': 'agCheckboxCellRenderer', 'cellEditor': 'agCheckboxCellEditor',},
    {'field': 'Belongs To Item Folder',},
]

headerName = [x['field'] for x in columns]
hiddenHeaders = ['Dine In Price','Bar Price','Pick Up Price','Take Out Price','Delivery Price','Open Price Item','Tax 2','Tax 3','This Is A Bar Item','This Is A Weighted Item','Tare','Barcode']
rows = []





#------------------------Start of Page------------------------------------------------

with ui.grid(columns=5): # Menu Item Input
    it = ui.input(label="Item Type", value="SM", validation={'field required': lambda val: len(val) > 0})

    ip = ui.number(label="Prices", value = 1, validation={'field required': lambda val: val > -1})

    fn = ui.input(label='Folder Name', value="Hot Coffee")

    ig = ui.input(label='Item Group', value="Coffee", validation={'field required': lambda val: len(val) > 0})

    ic = ui.input(label='Item Category', value="Drinks", validation={'field required': lambda val: len(val) > 0})



ui.button('Add row', icon='add', on_click=add)

grid = ui.aggrid({
    'columnDefs': columns,
    'defaultColDef': {'wrapHeaderText': True,'autoHeaderHeight': True, 'editable': True},
    'rowData': rows,
    'stopEditingWhenCellsLoseFocus': True
})

ui.button('Export', on_click=export)
ui.button('Terminal print', on_click=show_items)
ui.button('Show hidden', on_click=lambda: grid.run_grid_method('setColumnsVisible', hiddenHeaders, True))
ui.button('Hide hidden', on_click=lambda: grid.run_grid_method('setColumnsVisible', hiddenHeaders, False))

ui.run()