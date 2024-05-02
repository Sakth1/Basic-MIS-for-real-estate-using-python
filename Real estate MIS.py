#NAME: Sakthi Murugan C
#ROLL NO: 23121127


import pandas as pd
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from reportlab.platypus import SimpleDocTemplate, Paragraph
from reportlab.lib.styles import getSampleStyleSheet

class Property:
    def __init__(self, file_path):
        self.file_path = file_path
        self.properties = pd.read_excel(file_path)

    def display_properties(self):
        print(self.properties)

    def add_property(self):
        # Take user input
        prop_id = int(input("Enter property ID: "))
        address = input("Enter property address: ")
        price = int(input("Enter property price: "))
        status = input("Enter property status (available/rented): ")

        # Create a new DataFrame for the new data
        new_data = pd.DataFrame({
            'Property ID': [prop_id],
            'Address': [address],
            'Price': [price],
            'Status': [status]
        })

        # Append the new data to the existing DataFrame
        self.properties = pd.concat([self.properties, new_data], ignore_index=True)

        # Write the updated DataFrame back to the properties Excel file
        self.properties.to_excel(self.file_path, index=False)
        print(self.properties)

    def update_property(self):
        # Displaying existing properties
        print(self.properties)

        # Asking the user for which data to be updated
        while True:
            ch = input("Is the data to be changed a number?(Y/N) ")
            if ch == "Y":
                rw = int(input("Enter the row: "))
                col = input("Enter the column name: ")
                new_value = input("Enter the New value: ")
                new_value = int(new_value)
                self.properties.loc[rw, col] = new_value
                print("Replacement successful")
                break
            elif ch == "N":
                rw = int(input("Enter the row: "))
                col = input("Enter the column name: ")
                new_value = input("Enter the New value: ")
                self.properties.loc[rw, col] = new_value
                print("Replacement successful")
                break
            else:
                print("Enter a valid choice!")

        self.properties.to_excel(self.file_path, index=False)

class Client:
    def __init__(self, file_path):
        self.file_path = file_path
        self.clients = pd.read_excel(file_path)

    def display_clients(self):
        print(self.clients)

    def add_client(self):
        cli_id = int(input("Enter client ID: "))
        name = input("Enter client Name: ")
        rent = int(input("Enter client rent: "))

        # Create a new DataFrame for the new data
        new_data = pd.DataFrame({
            'Client ID': [cli_id],
            'Name': [name],
            'Rent': [rent]
        })

        # Append the new data to the existing DataFrame
        self.clients = pd.concat([self.clients, new_data], ignore_index=True)

        # Write the updated DataFrame back to the Excel file
        self.clients.to_excel(self.file_path, index=False)

def generate_report(properties, clients):
    # Creating a PDF file
    pdf_filename = "E:\\Users\\Desktop\\Real-Estate\\report.pdf"
    doc = SimpleDocTemplate(pdf_filename, pagesize=letter)
    styles = getSampleStyleSheet()

    elements = []

    # Report of available properties
    elements.append(Paragraph("Available Properties Report", styles['Heading1']))
    available_properties = properties.properties[properties.properties['Status'] == 'available']
    for index, row in available_properties.iterrows():
        property_info = "Property ID: {}, Address: {}, Price: {}".format(row['Property ID'], row['Address'], row['Price'])
        elements.append(Paragraph(property_info, styles['Normal']))

    # Report of clients
    elements.append(Paragraph("Clients Report", styles['Heading1']))
    for index, row in clients.clients.iterrows():
        client_info = "Client ID: {}, Name: {}, Rent: {}".format(row['Client ID'], row['Name'], row['Rent'])
        elements.append(Paragraph(client_info, styles['Normal']))

    # Income report
    rent = clients.clients["Rent"]
    income = sum(rent)
    elements.append(Paragraph(f"Total Income: {income}", styles['Heading3']))

    # Build PDF
    doc.build(elements)

def _main_():
    properties_file_path = "E:\\Users\\Desktop\\Real-Estate\\properties.xlsx"
    clients_file_path = "E:\\Users\\Desktop\\Real-Estate\\clients.xlsx"

    properties_obj = Property(properties_file_path)
    clients_obj = Client(clients_file_path)

    while True:
        print("Choose one option from below:")
        print("1. Display Every Property")
        print("2. Add Property")
        print("3. Display Every Client")
        print("4. Add Client")
        print("5. Update Property")
        print("6. Generate Report")
        print("7. End")
        choice = int(input("Enter Your Choice: "))
        if choice <= 7 and choice >= 1:
            if choice == 1:
                properties_obj.display_properties()
            elif choice == 2:
                properties_obj.add_property()
            elif choice == 3:
                clients_obj.display_clients()
            elif choice == 4:
                clients_obj.add_client()
            elif choice == 5:
                properties_obj.update_property()
            elif choice == 6:
                generate_report(properties_obj, clients_obj)
            elif choice == 7:
                break
        else:
            print("Enter a valid choice")


_main_()