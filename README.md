# RPA Data Management Automation

This project is an RPA (Robotic Process Automation) solution designed to automate the process of managing data requests in a SharePoint environment. The automation reads data from an Excel file stored in SharePoint, processes the data, and updates the Excel file with the results. The project is built using Python and leverages various libraries to interact with SharePoint and handle data processing.

## API Used

- **SharePoint API**: Used to interact with SharePoint for reading and writing data.
- **Microsoft Graph API**: Used for authentication and accessing SharePoint resources.

## Project Setup

### Extracting the Zip File

1. Download the project zip file.
2. Extract the zip file to your desired directory.

### Creating a Virtual Environment

1. Open a terminal and navigate to the project directory. (`src`)
2. Create a virtual environment in the project directory:
   ```sh
   python -m venv venv
   ```
3. Activate the virtual environment:
   - On macOS and Linux:
     ```sh
     source src/venv/bin/activate
     ```
   - On Windows:
     ```sh
     src\venv\Scripts\activate
     ```

### Download Dependencies

1. Install the required dependencies using `pip`:
   ```sh
   pip install python-dotenv msal requests
   ```

### Set the Configuration in `.env`

1. Create a `.env` file in the project directory.
2. Add the following configuration variables to the `.env` file:
   ```env
   CLIENT_ID=your_client_id
   TENANT_ID=your_tenant_id
   REDIRECT_URI=your_redirect_uri
   FREQUENCY=your_frequency
   RPABAY_DATA_MANAGEMENT_REQUEST_FORM=your_excel_name
   RPABAY_DATA_MANAGEMENT_REQUEST_FORM_SHEET=your_sheet_name
   RPABAY_DATA_MANAGEMENT=your_sharepoint_directory_rpabay
   RPABAY_DATA_GIDEN=your_sharepoint_directory_sent
   ```

### Register Application and Set Permissions in Azure

1. Go to the [Azure Portal](https://portal.azure.com/).
2. Navigate to **Azure Active Directory** > **App registrations** > **New registration**.
3. Register a new application and note down the **Application (client) ID** and **Directory (tenant) ID**.
4. Under **Manage**, select **API permissions** > **Add a permission**.
5. Add the following Microsoft Graph API permissions:
   - `User.Read`
   - `Files.ReadWrite.All`
   - `Sites.ReadWrite.All`
   - `Sites.Manage.All`
   - `Mail.Send`
6. Grant admin consent for the added permissions.

### Set the Configuration on Power Automate Flow

1. Open Power Automate.
2. Create a new flow or edit an existing flow.
3. Set the necessary configurations to match the variables in your `.env` file.
4. Ensure that the flow triggers and actions are correctly set up to interact with your SharePoint and other services as needed.

### Running the Project

1. Ensure the virtual environment is activated.
2. Run the main script:
   ```sh
   python main.py
   ```

This will start the process, reading the Excel file from SharePoint, processing the data, and performing the necessary actions as defined in your scripts.
