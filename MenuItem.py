class MenuItem:
    def __init__(self, name_type=None, group=None, category=None, price=None, folder=None):
        self.holder = []
        self.count = 0

        if name_type is not None:
            self.add_item(name_type, group, category, price, folder)
            

    def add_item(self, name_type:str, group:str, category:str, price:int, folder_name:str, print_loc:str):
        item = {
            'id' : self.count,
            'Menu Item Full Name': name_type,
            'Menu Item Group': group,
            'Menu Item Category': category,
            'Default Price': price,
            'Dine In Price': None,
            'Bar Price': None,
            'Pick Up Price': None,
            'Take Out Price': None,
            'Delivery Price': None,
            'Open Price Item': None,
            'POS Orders Print At': print_loc,
            'Tax 1': False,
            'Tax 2': False,
            'Tax 3': False,
            'This Is A Bar Item': False,
            'This Is A Weighted Item': False,
            'Tare': None,
            'Barcode': None,
            'Item Folder': False,
            'Belongs To Item Folder': folder_name,
        }
        self.holder.append(item)
        self.count += 1
        #print(self.holder)


    def add_dict(self, data):
        self.holder.append(data)


    def delete_item(self, item1):
        self.holder.remove(item1)
        #print(self.holder)

    def update_index(self, value):
        self.count = value

    def get_rows(self):
        return self.holder

    def get_index(self, item):
        return self.holder[item]['id'] if item in self.holder else -1