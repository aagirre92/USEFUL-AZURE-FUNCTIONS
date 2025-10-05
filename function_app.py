import azure.functions as func
import random
import json
import zeep
import base64
import xlrd
import logging

app = func.FunctionApp(http_auth_level=func.AuthLevel.FUNCTION)


@app.route(route="http_trigger_ando")
def http_trigger_ando(req: func.HttpRequest) -> func.HttpResponse:

    logging.info('Python HTTP trigger function processed a request.')

    name = req.params.get('name')
    if not name:
        try:
            req_body = req.get_json()
        except ValueError:
            pass
        else:
            name = req_body.get('name')

    if name:
        return func.HttpResponse(f"Hello, {name}. This HTTP triggered function executed successfully.")
    else:
        return func.HttpResponse(
            "This HTTP triggered function executed successfully. Pass a name in the query string or in the request body for a personalized response.",
            status_code=200
        )


@app.route(route="soap_trigger_ando")
def soap_trigger_ando(req: func.HttpRequest) -> func.HttpResponse:
    logging.info('Python SOAP trigger function processed a request.')

    wsdl = 'http://www.dneonline.com/calculator.asmx?WSDL'
    client = zeep.Client(wsdl=wsdl)
    first_num = req.params.get('first_num')
    second_num = req.params.get('second_num')
    if not first_num or not second_num:
        return func.HttpResponse(
            "Please pass both 'first_num' and 'second_num' in the query string.",
            status_code=400
        )
    try:
        result = client.service.Add(intA=int(first_num), intB=int(second_num))
        return func.HttpResponse(f"The result of the SOAP Add operation is: {result}")
    except Exception as e:
        logging.error(f"Error calling SOAP service: {e}")
        return func.HttpResponse(
            f"Error calling SOAP service: {e}",
            status_code=500
        )


@app.route(route="check_workday_inventory_balance")
def check_workday_inventory(req: func.HttpRequest) -> func.HttpResponse:
    logging.info(
        'Python check_workday_inventory function processed a request.')

    # Here you would implement the logic to check workday inventory
    # For demonstration, let's assume we return a static response
    part_number = req.params.get('part_number')
    location = req.params.get('location')
    if not part_number or not location:
        return func.HttpResponse(
            "Please pass both 'part_number' and 'location' in the query string.",
            status_code=400
        )

    # Simulate inventory check logic
    logging.info(
        f"Checking inventory for part number: {part_number} at location: {location}")
    # In a real scenario, you would query a database or an external service here
    # For now, we just log and return a success message

    if random.randint(1, 100) > 30:
        return_body = {
            "part_number": part_number,
            "location": location,
            "status": "Out of Stock",
            "quantity": 0
        }
    else:
        return_body = {
            "part_number": part_number,
            "location": location,
            "status": "In Stock",
            "quantity": random.randint(1, 100)
        }

    return func.HttpResponse(
        json.dumps(return_body),
        status_code=200
    )


@app.route(route="get_PO_xls_details")
def get_PO_xls_details(req: func.HttpRequest) -> func.HttpResponse:
    """
    This function receives a base64 encoded xls file and returns the PO details.
    The function returns the vendor id (cell A1), vendor name (cell A20), PO number (cell K18) in json format.
    """
    logging.info('Python get_PO_xls_details function processed a request.')

    # Here you would implement the logic to get PO XLS details
    # For demonstration, let's assume we return a static response
    if req.method != "POST":
        logging.error("Invalid request method. Only POST is allowed.")
        return func.HttpResponse(
            body=json.dumps({
                "error": "Invalid request method. Only POST is allowed."
            }),
            mimetype="application/json",
            status_code=400
        )
    try:
        req_body = req.get_json()
        xls_base64 = req_body.get('xls_base64')
    except Exception as e:
        logging.error(f"Error parsing request body: {e}")
        return func.HttpResponse(
            body=json.dumps({
                "error": f"Error parsing request body: {e}"
            }),
            mimetype="application/json",
            status_code=400
        )
    if not xls_base64:
        logging.error("Missing 'xls_base64' in request body.")
        return func.HttpResponse(
            body=json.dumps({
                "error": "Missing 'xls_base64' in request body."
            }),
            mimetype="application/json",
            status_code=400
        )
    try:
        xls_data = base64.b64decode(xls_base64)
    except Exception as e:
        logging.error(f"Error decoding base64 string: {e}")
        return func.HttpResponse(
            body=json.dumps({
                "error": f"Error decoding base64 string: {e}"
            }),  
            mimetype="application/json",          
            status_code=400
        )

    try:

        workbook = xlrd.open_workbook(file_contents=xls_data)
        sheet = workbook.sheet_by_index(0)
        vendor_id = sheet.cell_value(0, 0)  # Cell A1
        vendor_name = sheet.cell_value(19, 0)  # Cell A20
        po_number = sheet.cell_value(17, 10)  # Cell K18
        return_body = {
            "vendor_id": vendor_id,
            "vendor_name": vendor_name,
            "po_number": po_number
        }
        logging.info(f"Extracted PO details: {return_body}")
        return func.HttpResponse(
            body=json.dumps(return_body),
            mimetype="application/json",
            status_code=200
        )
    except Exception as e:
        logging.error(f"Error processing xls file: {e}")
        return func.HttpResponse(
            body=json.dumps({
                "error": f"Error processing xls file: {e}"
            }), 
            mimetype="application/json",
            status_code=500
        )
