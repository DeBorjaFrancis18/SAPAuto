from datetime import datetime, timedelta
from pyrfc import Connection
import openpyxl

# Read SAP credentials
def read_sap_credentials(filename):
    with open(filename, 'r') as file:
        lines = file.readlines()
        username = lines[0].strip().split(':')[1]
        password = lines[1].strip().split(':')[1]
    return username, password

# Create a sales order
def create_sales_order(conn, header_info, item_info):
    try:
        # Sales order header
        sales_order_header = {
            'SalesDocument': '',
            'SalesDocumentType': header_info['OrderType'],
            'SalesOrganization': header_info['SalesOrg'],
            'DistributionChannel': header_info['DistributionChannel'],
            'Division': header_info['Division'],
            'SoldToParty': header_info['SoldToParty'],
            'ShipToParty': header_info['ShipToParty'],
            'PurchaseOrderNumber': header_info['PONumber'],
            'PurchaseOrderDate': header_info['PODate'],  # Assuming PO Date is not provided in the Excel
            'DeliveryPlant': header_info['DeliveryPlant'],
            'RequestedDeliveryDate': header_info['DeliveryDate'],  # Assuming Delivery Date is not provided in the Excel
            'ShippingConditions': header_info['ShippingCondition'],
            'Vendor': header_info['Vendor'],
        }

        # Call BAPI to create sales order header
        sales_order_header_result = conn.call('BAPI_SALESORDER_CREATEFROMDAT2', SalesOrderHeaderIn=sales_order_header)

        # Sales order created
        sales_document = sales_order_header_result['SalesOrder']

        # Sales order items
        for item in item_info:
            sales_order_item = {
                'SalesDocument': sales_document,
                'SalesDocumentItem': '',
                'Material': item['Material'],
                'RequestedQuantity': item['Quantity'],
                'RequestedQuantityUnit': item['UoM'],
                'Plant': item['Sloc'],  # Assuming Storage Location is mapped to Plant
            }

            # Call BAPI to create sales order item
            sales_order_item_result = conn.call('BAPI_SALESORDER_CREATEFROMDATA', OrderItemsIn=[sales_order_item])

        print("Sales order created successfully.")
        return sales_document

    except Exception as e:
        print(f"Error creating sales order: {e}")
        raise

try:
    # SAP credentials
    username, password = read_sap_credentials("SAPCreds.txt")

    # Connect to SAP
    conn = Connection(user=username, passwd=password, ashost='your_sap_server', sysnr='00', client='100', lang='EN')

    # Read Header information from Excel
    wb = openpyxl.load_workbook('SAPAuto.xlsx')
    header_sheet = wb['Header']
    header_info = {
        'SAPBox': header_sheet['A2'].value,
        'OrderType': header_sheet['B2'].value,
        'SalesOrg': header_sheet['C2'].value,
        'DistributionChannel': header_sheet['D2'].value,
        'Division': header_sheet['E2'].value,
        'SoldToParty': header_sheet['F2'].value,
        'ShipToParty': header_sheet['G2'].value,
        'DeliveryPlant': header_sheet['H2'].value,
        'PONumber': 'TA_' + datetime.now().strftime("%m%d%y%H%M%S"),  # Generate PO Number
        'PODate': (datetime.now() - timedelta(days=1)).strftime("%m%d%y%H%M%S"),  # Subtract 1 day from current date
        'DeliveryDate': (datetime.now() + timedelta(days=int(header_sheet['I2'].value))).strftime("%m%d%y%H%M%S"),  # Calculate Delivery Date based on Transit Time
        'ShippingCondition': header_sheet['J2'].value,
        'Vendor': header_sheet['K2'].value,
    }

    # Read Item information from Excel
    item_sheet = wb['Item']
    item_info = []
    for row in item_sheet.iter_rows(min_row=2, values_only=True):
        item_info.append({
            'Material': row[0],
            'Quantity': row[1],
            'UoM': row[2],
            'Sloc': row[3]
        })

    # Sales Order in SAP
    sales_order = create_sales_order(conn, header_info, item_info)

except Exception as e:
    print(f"Error: {e}")
