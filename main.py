# import openpyxl
# from openpyxl import Workbook
import os
from kivymd.app import MDApp
from kivy.lang import Builder
from kivymd.uix.pickers import MDModalDatePicker
from kivymd.uix.menu import MDDropdownMenu
from kivymd.uix.navigationbar import MDNavigationBar, MDNavigationItem
from kivymd.uix.snackbar import MDSnackbar, MDSnackbarSupportingText
from kivy.properties import StringProperty
from kivy.metrics import dp
from kivymd.uix.label import MDLabel
from googleapiclient.discovery import build
from google.oauth2.service_account import Credentials
from datetime import datetime
from dotenv import load_dotenv

load_dot = load_dotenv()


SCOPES_UPLOAD = os.getenv("SCOPES_UPLOAD")
SERVICE_ACCOUNT_FILE = os.getenv("SERVICE_ACCOUNT_FILE")
SCOPES = os.getenv("SCOPES")
spreadsheet_id = os.getenv("spreadsheet_id")


class BaseMDNavigationItem(MDNavigationItem):
    icon = StringProperty()
    text = StringProperty()


KV_string = '''
<BaseMDNavigationItem>
    MDNavigationItemIcon:
        icon: root.icon
    MDNavigationItemLabel:
        text: root.text
    
MDScreen:
    MDBoxLayout:
        orientation: "vertical"
        md_bg_color: "white"

        MDScreenManager:
            id: screen_manager
            
            MDScreen:    
                name: "Add Expense" 
                MDBoxLayout:
                    orientation: "vertical"
                    md_bg_color: "white"

                    MDLabel:  
                        text: "Expense Tracker"
                        text_color: "blue"
                        halign: "center"
                        size_hint_y: None
                        height: dp(50)
                    ScrollView: 
                        MDBoxLayout:
                            md_bg_color: "white"
                            orientation: "vertical"
                            spacing: "15dp"
                            padding: "20dp"  
                            size_hint_y: None
                            height: self.minimum_height

                            MDLabel:
                                text: "Date"
                                halign: "center"
                                
                            MDTextField:
                                id: date_field
                                mode: "outlined"
                                pos_hint: {"center_x": 0.5, "center_y": 0.9}
                                size_hint_x: 0.9
                                on_focus: app.show_date_picker(self.focus)

                                MDTextFieldHintText:
                                    text: "Select Date YY-MM-DD"

                                MDTextFieldTrailingIcon:
                                    icon: "calendar"

                            MDLabel:
                                text: "Expense Type"
                                halign: "center"

                            MDTextField:
                                id: field
                                hint_text: "Select Expense Type"
                                pos_hint: {"center_x": 0.5, "center_y": 0.9}
                                size_hint_x: 0.9
                                on_focus: if self.focus: app.expense_menu.open()
                                MDTextFieldHintText:
                                    text: "Select Expense Type"

                            MDLabel:
                                text: "Sub-Type"
                                halign: "center"

                            MDTextField:
                                id: subtype_field
                                hint_text: "Select Sub-Type"
                                pos_hint: {"center_x": 0.5, "center_y": 0.8}
                                size_hint_x: 0.9
                                on_focus: if self.focus: app.subtype_menu.open()
                                MDTextFieldHintText:
                                    text: "Select Sub-Type"

                            MDLabel:
                                text: "Description"
                                halign: "center"

                            MDTextField:
                                id: description_field
                                mode: "outlined"
                                size_hint_x:  0.9
                                pos_hint: {"center_x": 0.5, "center_y": 0.5}
                                MDTextFieldHintText:
                                    text: "Description"


                            MDLabel:
                                text: "Amount"
                                halign: "center"

                            MDTextField:
                                id: amount_field
                                mode: "outlined"
                                input_filter: "int"
                                size_hint_x: 0.9
                                pos_hint: {"center_x": 0.5, "center_y": 0.5}
                                MDTextFieldHintText:
                                    text: "Enter amount"
                                MDTextFieldHelperText:
                                    text: "Enter amount"

                            MDButton:
                                style: "elevated"
                                style:"filled"
                                padding: "20dp"
                                theme_width: "Custom"
                                height: "56dp"
                                size_hint_x: .5
                                pos_hint: {"center_x": .5, "center_y": .5}
                                on_release: app.submit_expense()
                                MDButtonText:
                                    text: "Submit"
                                    pos_hint: {"center_x": .5, "center_y": .5}
                        
            MDScreen:
                name: "View Expense"

                MDBoxLayout:
                    orientation: "vertical"

                    MDLabel: 
                        text: "Expense Tracker"
                        text_color: "blue"
                        halign: "center"
                        size_hint_y: None
                        height: dp(50) 
                        md_bg_color: "white"
                    MDBoxLayout:
                        orientation: "horizontal"
                        md_bg_color: "white"
                        size_hint_y: None  
                        height: dp(70)  
                        padding: "4dp"  
                        spacing: "10dp"  

                        MDTextField:
                            id: month_dropdown
                            hint_text: "Select Month"
                            mode: "outlined"
                            readonly: True
                            text: app.selected_month  # Set default to current month
                            on_focus: if self.focus: app.month_menu.open()
                            size_hint_x: 0.2
                        MDTextField:
                            id: year_dropdown
                            hint_text: "Select Year"
                            mode: "outlined"
                            readonly: True
                            text: app.selected_year  
                            on_focus: if self.focus: app.year_menu.open()
                            size_hint_x: 0.2

                        MDButton:
                            on_release: app.display_expenses(app.root.ids.expense_list_grid)
                            size_hint_x: None
                            width: dp(100)  # Control the width of the button
                            height: dp(40)  # Control the height of the button
                            padding: "10dp"  # Reduce padding to make it more compact
                            MDButtonText:
                                text: "Submit"
                                
                    ScrollView:
                        do_scroll_x: False
                        do_scroll_y: True

                        MDBoxLayout:
                            md_bg_color: "white"
                            id: expense_list_viewlayout
                            orientation: "vertical"
                            size_hint_y: None 
                            height: self.minimum_height 
        
                            MDGridLayout:
                                id: expense_list_grid
                                cols: 5
                                # size_hint_x: None # it is not responsive (horizantal scroll should add )
                                size_hint_y: None
                                height: self.minimum_height
                                width: dp(800)  
                                row_default_height: dp(30)
                                row_force_default: True
                                # padding: "10dp"
                                # spacing: "10dp"
                            
                                
        MDNavigationBar:
            on_switch_tabs: app.on_switch_tabs(*args)
            BaseMDNavigationItem:
                icon: "plus"
                text: "Add Expense"
                active: True
            BaseMDNavigationItem:
                icon: "eye"
                text: "View Expense"
        
'''

class AddExpenseScreen(MDApp):
    def __init__(self, **kwargs):
        super().__init__(**kwargs)
    
    def on_switch_tabs(
        self,
        bar: MDNavigationBar, 
        item: MDNavigationItem,
        item_icon: str,
        item_text: str,
    ):
        self.root.ids.screen_manager.current = item_text
        if item_text == "View Expense":
            self.load_expenses()

    def build(self):
        self.theme_cls.primary_palette = "Orange"
        self.theme_cls.theme_style = "Light"
        current_date = datetime.now()
        self.current_month = current_date.strftime('%B')
        self.current_year = current_date.year
        self.selected_month = self.current_month
        self.selected_year = str(self.current_year)
        # Initialize dropdown menus
        self.month_menu = None  
        self.year_menu = None 
        # Loading the KV 
        self.screen = Builder.load_string(KV_string)
        # Expense Type Dropdown Menu
        expense_list = ["Labour", "Material", "Others"]
        self.expense_menu_items = [
            {"text": i, "on_release": lambda x=i: self.set_expense_type(x)} for i in expense_list
        ]
        self.expense_menu = MDDropdownMenu(
            caller=self.screen.ids.field, items=self.expense_menu_items, position="bottom"
        )
        # Sub-Type Dropdown Menu
        self.subtype_menu_items = []
        self.subtype_menu = MDDropdownMenu(
            caller=self.screen.ids.subtype_field, items=self.subtype_menu_items, position="bottom"
        )
        # Month Dropdown
        months = [
            "January", "February", "March", "April", "May", "June",
            "July", "August", "September", "October", "November", "December"
        ]
        self.month_menu = MDDropdownMenu(
            caller=self.screen.ids.month_dropdown,
            items=[
                {"text": m, "on_release": lambda x=m: self.set_month(x)}
                for m in months
            ],
            width_mult=4,
        )
        # Year Dropdown (Current Year and Last 10 Years)
        current_year = datetime.now().year
        years = [str(year) for year in range(current_year, current_year - 11, -1)]
        self.year_menu = MDDropdownMenu(
            caller=self.screen.ids.year_dropdown,
            items=[
                {"text": y, "on_release": lambda x=y: self.set_year(x)}
                for y in years
            ],
            width_mult=4,
        )
        # Set default month and year
        self.screen.ids.month_dropdown.text = months[datetime.now().month - 1]
        self.screen.ids.year_dropdown.text = str(current_year)
        return self.screen

    def show_date_picker(self,focus):
        if not focus:
            return
        date_dialog = MDModalDatePicker()
        date_dialog.open()
        date_dialog.bind(on_ok=self.on_ok, on_cancel=self.on_cancel)
        
    def on_ok(self, instance_date_picker):
        # print(instance_date_picker.get_date()[0])
        self.screen.ids.date_field.text = str(instance_date_picker.get_date()[0])
        instance_date_picker.dismiss()

    def on_cancel(self,instance_date_picker,*args):
        self.screen.ids.date_field.text = "No date selected" 
        instance_date_picker.dismiss()

    def set_expense_type(self, expense_type):
        self.screen.ids.field.text = expense_type
        self.expense_menu.dismiss()
        self.screen.ids.subtype_field.text = ""
        # Update sub-type menu based on the selected expense type
        if expense_type == "Labour":
            subtype_list = [
                "Wood work", "Plumbing work", "Tiles work",
                "Construction work", "Electricity work", "Others"
            ]
        elif expense_type == "Material":
            subtype_list = ["Cement", "Steel", "Wood", "Electricals", "Plumbing", "Others"]
        else:
            subtype_list = ["Others"]
        # Update Sub-Type Menu
        self.subtype_menu_items = [
            {"text": i, "on_release": lambda x=i: self.set_subtype(x)} for i in subtype_list
        ]
        self.subtype_menu.items = self.subtype_menu_items

    def set_subtype(self, subtype):
        self.screen.ids.subtype_field.text = subtype
        self.subtype_menu.dismiss()

    def submit_expense(self):
        date = self.screen.ids.date_field.text
        expense_type = self.screen.ids.field.text
        subtype = self.screen.ids.subtype_field.text
        description = self.screen.ids.description_field.text
        amount = self.screen.ids.amount_field.text

         # Validate inputs
        if not date or date == "No date selected":
            self.show_message("Please select a date.")
            return
        if not expense_type:
            self.show_message("Please select an expense type.")
            return
        if not subtype:
            self.show_message("Please select a subtype.")
            return
        if not description:
            self.show_message("Please enter a description.")
            return
        if not amount or not amount.isdigit():
            self.show_message("Please enter a valid amount.")
            return
        # if not all([date, expense_type, subtype, description, amount]):
        #     message = "Please fill all fields!"
        #     self.show_message( message)
        #     return
        # ----------local file saving data -------------
        # file_path = "expenses.xlsx"
        # if not os.path.exists(file_path):
        #     workbook = Workbook()
        #     sheet = workbook.active
        #     sheet.append(["Date", "Expense Type", "Sub-Type", "Description", "Amount"])
        #     workbook.save(file_path)
        # workbook = openpyxl.load_workbook(file_path)
        # sheet = workbook.active
        # sheet.append([date, expense_type, subtype, description, amount])
        # workbook.save(file_path)
        # ---------------------------------------
        self.upload_to_google_sheets(date, expense_type,subtype, description, amount,SCOPES_UPLOAD,SERVICE_ACCOUNT_FILE,spreadsheet_id)
        message_success = "Expense saved successfully!"
        self.show_message( message_success)
        self.clear_fields()

    def clear_fields(self):
        self.screen.ids.date_field.text = ""
        self.screen.ids.field.text = ""
        self.screen.ids.subtype_field.text = ""
        self.screen.ids.description_field.text = ""
        self.screen.ids.amount_field.text = ""

    def upload_data_to_google_workspace(self):
        pass

    def show_message(self, message):
        MDSnackbar(
            MDSnackbarSupportingText(
                text=message,
                size_hint_x=0.8,
                pos_hint={"center_x": 0.5, "center_y": 0.5},    
            ),
            orientation="horizontal",
            pos_hint={"center_x": 0.9, "center_y": 0.2},
            size_hint=(0.4, 0.4),
            duration=1,
        ).open()
    
    def load_expenses(self):
        print(SCOPES,SERVICE_ACCOUNT_FILE,spreadsheet_id)
        grid = self.screen.ids.expense_list_grid
        grid.clear_widgets()
        # Add Headers
        # headers = ["Date", "Type", "Sub-Type", "Description", "Amount"]
        # for header in headers:
        #     grid.add_widget(MDLabel(text=header, bold=True, halign="center"))
        # Populate Expenses
        self.display_expenses(grid,SCOPES,SERVICE_ACCOUNT_FILE,spreadsheet_id)

    def upload_to_google_sheets(self, date, expense_type, subtype, description, amount,SCOPES_UPLOAD,SERVICE_ACCOUNT_FILE,spreadsheet_id):
        SCOPES = [SCOPES_UPLOAD]       
        SERVICE_ACCOUNT_FILE = SERVICE_ACCOUNT_FILE
        creds = Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=SCOPES)
        service = build('sheets', 'v4', credentials=creds)
        spreadsheet_id = spreadsheet_id
        # Determine the month sheet name (e.g., "January 2025")
        date_obj = datetime.strptime(date, '%Y-%m-%d')
        month_sheet_name = date_obj.strftime('%B %Y')
        print(month_sheet_name)
        # Check if the sheet for the month exists
        sheets = service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute().get('sheets', [])
        sheet_names = [sheet['properties']['title'] for sheet in sheets]
        print(sheet_names)
        if month_sheet_name not in sheet_names:
            # Create the monthly sheet
            service.spreadsheets().batchUpdate(
                spreadsheetId=spreadsheet_id,
                body={
                    "requests": [
                        {
                            "addSheet": {
                                "properties": {
                                    "title": month_sheet_name
                                }
                            }
                        }
                    ]
                }
            ).execute()
            # Add headers to the new sheet
            header_values = [["Date", "Expense Type","subtype", "Description", "Amount"]]
            service.spreadsheets().values().update(
                spreadsheetId=spreadsheet_id,
                range=f"{month_sheet_name}!A1",
                valueInputOption="USER_ENTERED",
                body={"values": header_values}
            ).execute()
        # Append the new expense data to the corresponding month's sheet
        values = [[date, expense_type,subtype, description, amount]]
        service.spreadsheets().values().append(
            spreadsheetId=spreadsheet_id,
            range=f"{month_sheet_name}!A1",
            valueInputOption='USER_ENTERED',
            insertDataOption='INSERT_ROWS',
            body={'values': values}
        ).execute()
# -------------------------------------------------------------------------------------------------------------------4
    def set_month(self, month):
        self.selected_month = month
        self.screen.ids.month_dropdown.text = month
        self.month_menu.dismiss()

    def set_year(self, year):
        self.selected_year = year
        self.screen.ids.year_dropdown.text = year
        self.year_menu.dismiss()
       
    def clear_grid(self, grid):
        # Remove all widgets in the grid
        grid.clear_widgets()

    def display_expenses(self, grid,SCOPES,SERVICE_ACCOUNT_FILE,spreadsheet_id ):
        print(self.selected_month,self.selected_year)
        self.clear_grid(grid)
        SCOPES = [SCOPES]
        SERVICE_ACCOUNT_FILE = SERVICE_ACCOUNT_FILE
        # Authenticate using service account credentials
        creds = Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=SCOPES)
        service = build('sheets', 'v4', credentials=creds)
        # The ID of your Google Sheets spreadsheet
        spreadsheet_id = spreadsheet_id
        # Get the current month and year
        # current_month_year = datetime.now().strftime('%B %Y')
        current_month_year = f"{self.selected_month} {self.selected_year}"
        print(current_month_year)
        try:
            sheets = service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute().get('sheets', [])
            sheet_names = [sheet['properties']['title'] for sheet in sheets]
            # Check if the sheet for the current month exists
            if current_month_year in sheet_names:
                # Fetch the data from the current month's sheet
                range_name = f"{current_month_year}!A2:E"  # Skip header row
                result = service.spreadsheets().values().get(
                    spreadsheetId=spreadsheet_id,
                    range=range_name
                ).execute()
                values = result.get('values', [])
                
                if not values:    
                    self.show_message2(f"No expenses recorded for this {current_month_year}.")
                    return
                # Display the fetched data in the grid
                grid = self.screen.ids.expense_list_grid
                grid.clear_widgets()
                # --------------------------------------------------------------
                
                total_amount = 0.0  
                for row in values:
                    try:
                        # Assuming the amount is in the 5th column (index 4)
                        amount = float(row[4]) if row[4] else 0.0 
                        total_amount += amount 
                    except ValueError:
                        continue  

                # grid.add_widget(MDLabel(text=str(total_amount) if value else "", halign="center"))
                total_label = MDLabel(
                    text=f"Total: {total_amount:.2f}",  
                    halign="center",
                    size_hint_y=None,
                    height="50dp",  
                    md_bg_color="lightgreen", 
                    padding="10dp",  
                )
                self.screen.ids.expense_list_viewlayout.add_widget(total_label) 
                # ----------------------------------------------------------------------------
                # Add Headers
                headers = ["Date", "Type", "Sub-Type", "Description", "Amount"]
                for header in headers:
                    grid.add_widget(MDLabel(text=header, bold=True, halign="center"))
                for row in values:
                    for value in row:
                        # print(value)
                        grid.add_widget(MDLabel(text=str(value) if value else "", halign="center"))
  
            else:
                self.show_message2(f"No sheet found for {current_month_year}.")
        except Exception as e:
            self.show_message(f"Error fetching data: {str(e)}")
    def show_message2(self, message):
        MDSnackbar(
            MDSnackbarSupportingText(
                text=message,
                size_hint_x=0.5,
                pos_hint={"center_x": 0.5, "center_y": 0.5},    
            ),
            orientation="horizontal",
            pos_hint={"center_x": 0.5, "center_y": 0.5},
            size_hint=(0.4, 0.4),
            duration=1,
        ).open()
    

if __name__ == "__main__":
    AddExpenseScreen().run()
