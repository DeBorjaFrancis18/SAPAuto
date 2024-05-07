from pyrfc import Connection
import openpyxl
from datetime import datetime, timedelta

def read_excel_data(file_path):
    try:
        wb = openpyxl.load_workbook(file_path)
        header_sheet = wb['Header']
        item_sheet = wb['Item']
        
        header_info = {
            'SAPBox': header_sheet['A2'].value,
            'OrderType': header_sheet['B2'].value,
            'SalesOrg': header_sheet['C2'].value,
            'DistributionChannel': header_sheet['D2'].value,
            'Division': header_sheet['E2'].value,
            'SoldToParty': header_sheet['F2'].value,
            'ShipToParty': header_sheet['G2'].value,
            'DeliveryPlant': header_sheet['H2'].value,
            'PONumber': 'TA_' + datetime.now().strftime("%Y%m%d%H%M%S"),
            'PODate': datetime.now().strftime("%Y%m%d"),
            'DeliveryDate': (datetime.now() + timedelta(days=int(header_sheet['I2'].value))).strftime("%Y%m%d"),
            'ShippingCondition': header_sheet['J2'].value,
            'Vendor': header_sheet['K2'].value,
        }

        item_info = []
        for row in item_sheet.iter_rows(min_row=2, values_only=True):
            item_info.append({
                'Material': row[0],
                'Quantity': row[1],
                'UoM': row[2],
                'Sloc': row[3]
            })

        return header_info, item_info
    except Exception as e:
        raise Exception(f"Error reading Excel data: {e}")

def create_sales_order(conn, header_info, item_info):
    try:
        sales_order_header = {
            'SalesDocument': '',
            'SalesDocumentType': header_info['OrderType'],
            'SalesOrganization': header_info['SalesOrg'],
            'DistributionChannel': header_info['DistributionChannel'],
            'Division': header_info['Division'],
            'SoldToParty': header_info['SoldToParty'],
            'ShipToParty': header_info['ShipToParty'],
            'PurchaseOrderNumber': header_info['PONumber'],
            'PurchaseOrderDate': header_info['PODate'],
            'DeliveryPlant': header_info['DeliveryPlant'],
            'RequestedDeliveryDate': header_info['DeliveryDate'],
            'ShippingConditions': header_info['ShippingCondition'],
            'Vendor': header_info['Vendor'],
        }

        sales_order_header_result = conn.call('BAPI_SALESORDER_CREATEFROMDAT2', SalesOrderHeaderIn=sales_order_header)

        if sales_order_header_result['Return'] and sales_order_header_result['Return'][0]['Type'] == 'E':
            raise Exception(f"Error creating sales order header: {sales_order_header_result['Return'][0]['Message']}")

        sales_document = sales_order_header_result['SalesOrder']

        for item in item_info:
            sales_order_item = {
                'SalesDocument': sales_document,
                'SalesDocumentItem': '',
                'Material': item['Material'],
                'RequestedQuantity': item['Quantity'],
                'RequestedQuantityUnit': item['UoM'],
                'Plant': item['Sloc'],
            }

            sales_order_item_result = conn.call('BAPI_SALESORDER_CREATEFROMDATA', OrderItemsIn=[sales_order_item])

            if sales_order_item_result['Return'] and sales_order_item_result['Return'][0]['Type'] == 'E':
                raise Exception(f"Error creating sales order item: {sales_order_item_result['Return'][0]['Message']}")

        return sales_document
    except Exception as e:
        raise Exception(f"Error creating sales order: {e}")

try:
    # Assume conn is already established
    header_info, item_info = read_excel_data('SAPAuto.xlsx')
    sales_order = create_sales_order(conn, header_info, item_info)
    print(f"Sales order created successfully. Sales Document: {sales_order}")
except Exception as e:
    print(f"Error: {e}")
