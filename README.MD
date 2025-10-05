# Azure Functions - Andoni 101

This repository contains a collection of Azure Functions written in Python that demonstrate various integration patterns including HTTP triggers, SOAP service consumption, inventory management, and Excel file processing.

## Overview

This Azure Functions app includes four main endpoints designed to showcase different enterprise integration scenarios:

1. **HTTP Trigger** - Basic HTTP endpoint with personalized responses
2. **SOAP Service Integration** - Demonstrates calling external SOAP web services
3. **Inventory Management** - Mock inventory balance checking system
4. **Excel File Processing** - Processes base64-encoded Excel files to extract PO details

## Prerequisites

- Python 3.8 or higher
- Azure Functions Core Tools
- Azure CLI (for deployment)

## Dependencies

The project uses the following Python packages:
- `azure-functions` - Azure Functions runtime for Python
- `zeep` - Python SOAP client library
- `xlrd` - Excel file reading library

## Local Development

### Setup

1. Clone the repository
2. Create a virtual environment:
   ```bash
   python -m venv venv
   ```
3. Activate the virtual environment:
   ```bash
   # Windows
   venv\Scripts\activate
   
   # Linux/Mac
   source venv/bin/activate
   ```
4. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```

### Running Locally

Start the Azure Functions runtime:
```bash
func start
```

The functions will be available at `http://localhost:7071/api/`

## API Endpoints

### 1. HTTP Trigger Ando

**Endpoint:** `GET/POST /api/http_trigger_ando`

A basic HTTP triggered function that returns a personalized greeting.

**Parameters:**
- `name` (optional) - Name for personalized response (query parameter or JSON body)

**Example Requests:**
```bash
# Query parameter
curl "http://localhost:7071/api/http_trigger_ando?name=John"

# JSON body
curl -X POST "http://localhost:7071/api/http_trigger_ando" \
  -H "Content-Type: application/json" \
  -d '{"name": "John"}'
```

**Response:**
```
Hello, John. This HTTP triggered function executed successfully.
```

### 2. SOAP Trigger Ando

**Endpoint:** `GET /api/soap_trigger_ando`

Demonstrates SOAP service integration by calling an external calculator web service.

**Parameters:**
- `first_num` (required) - First number for addition
- `second_num` (required) - Second number for addition

**Example Request:**
```bash
curl "http://localhost:7071/api/soap_trigger_ando?first_num=5&second_num=3"
```

**Response:**
```
The result of the SOAP Add operation is: 8
```

### 3. Check Workday Inventory Balance

**Endpoint:** `GET /api/check_workday_inventory_balance`

Mock inventory management system that simulates checking inventory balance for parts.

**Parameters:**
- `part_number` (required) - Part number to check
- `location` (required) - Location to check inventory

**Example Request:**
```bash
curl "http://localhost:7071/api/check_workday_inventory_balance?part_number=ABC123&location=Warehouse1"
```

**Response:**
```json
{
  "part_number": "ABC123",
  "location": "Warehouse1",
  "status": "In Stock",
  "quantity": 45
}
```

### 4. Get PO XLS Details

**Endpoint:** `POST /api/get_PO_xls_details`

Processes base64-encoded Excel files to extract Purchase Order details from specific cells.

**Request Body:**
```json
{
  "xls_base64": "UEsDBBQAAAAIAC..."
}
```

**Extracted Data:**
- Vendor ID from cell A1
- Vendor Name from cell A20
- PO Number from cell K18

**Example Request:**
```bash
curl -X POST "http://localhost:7071/api/get_PO_xls_details" \
  -H "Content-Type: application/json" \
  -d '{"xls_base64": "your_base64_encoded_excel_file_here"}'
```

**Response:**
```json
{
  "vendor_id": "VENDOR001",
  "vendor_name": "ABC Supply Company",
  "po_number": "PO-2024-001"
}
```

## Error Handling

All endpoints include comprehensive error handling:
- Input validation
- Proper HTTP status codes
- Detailed error messages in JSON format
- Logging for debugging purposes

## Configuration

### Authentication

The functions are configured with `AuthLevel.FUNCTION`, requiring a function key for access when deployed to Azure.

### Local Settings

Configure `local.settings.json` for local development (this file is not included in the repository for security):

```json
{
  "IsEncrypted": false,
  "Values": {
    "AzureWebJobsStorage": "UseDevelopmentStorage=true",
    "FUNCTIONS_WORKER_RUNTIME": "python"
  }
}
```

## Deployment

To deploy to Azure:

1. Create an Azure Function App
2. Deploy using Azure Functions Core Tools:
   ```bash
   func azure functionapp publish <function-app-name>
   ```

## Testing

Each function includes comprehensive logging and error handling. Monitor function execution through:
- Local console output during development
- Azure Application Insights when deployed
- Azure Function App logs in the Azure portal

## Contributing

When adding new functions:
1. Follow the existing code structure
2. Include proper error handling and logging
3. Update this README with endpoint documentation
4. Add any new dependencies to `requirements.txt`

## License

This project is for educational and demonstration purposes.