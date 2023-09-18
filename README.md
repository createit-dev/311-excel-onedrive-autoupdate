# Updating Excel Files on OneDrive with Node.js and Microsoft Graph

This script provides a solution to automate the process of updating Excel sheets on OneDrive with data from an external API. Built using Node.js, the script leverages `axios` for HTTP requests and `ExcelJS` for manipulating Excel sheets.

## Features

1. **OAuth Authentication**: The script authenticates with Microsoft Graph API using OAuth to get the access token, allowing it to interact with OneDrive files.
2. **File Download**: Downloads the specified Excel file from OneDrive.
3. **Excel Manipulation**: The script reads, updates, and adds rows to the downloaded Excel sheet based on the data fetched from an external API.
4. **File Upload**: After updating the Excel sheet, the script uploads the updated file back to OneDrive.

## Requirements
For the script to function correctly, it is imperative to use **OneDrive for Business** and not the regular OneDrive.  OneDrive for Business is closely integrated with SharePoint, allowing for seamless sharing and collaboration on files stored on SharePoint sites. Our script specifically targets SharePoint document libraries, making this integration crucial.

The script relies on the Microsoft Graph API for accessing and modifying files on OneDrive. OneDrive for Business offers broader and more advanced capabilities when accessed via the Graph API compared to the regular OneDrive.

## Use-Cases for the Excel Automation Script

The script's flexibility to fetch data from external sources and update Excel sheets on OneDrive for Business presents a wide range of potential use-cases. Here are five scenarios where this script can be invaluable:

### 1. **HR Recruitment Tracker**:
Organizations can use the script to automatically update a centralized Excel sheet with candidate details fetched from various job portals or recruitment platforms. As new candidates apply or existing candidates update their profiles, the Excel sheet on OneDrive remains up-to-date, providing HR teams with real-time insights.

### 2. **Sales Pipeline Management**:
Sales teams can use the script to auto-update an Excel sheet with the latest lead and customer data fetched from CRM systems like Salesforce or HubSpot. This ensures that sales managers always have the most recent snapshot of the sales pipeline, enabling better forecasting and resource allocation.

### 3. **Inventory Management**:
Retailers and e-commerce platforms can utilize the script to automatically update stock levels in a centralized Excel sheet. By fetching data from warehouse management systems or point-of-sale systems, businesses can maintain real-time visibility into inventory levels, helping prevent stockouts or overstock situations.

### 4. **Event Registration**:
For event organizers or training institutions, the script can be used to update an Excel sheet with attendee or participant details fetched from event registration platforms. This aids in efficient event management, ensuring all stakeholders have access to the latest registration data.

### 5. **Financial Reporting**:
Finance teams can employ the script to periodically update financial reports in an Excel sheet based on data fetched from accounting software or ERP systems. By automating this process, businesses can ensure that financial stakeholders always have access to the most recent financial data, facilitating timely decision-making.

---

## Prerequisites

Before running the script, you need to:

1. Azure Application Registration: Ensure you have registered an application within the Azure portal. This process provides you with the necessary credentials, such as AZURE_CLIENT_ID and AZURE_CLIENT_SECRET, which are vital for the script's authentication with Azure services.
2. Azure App Permission Setup: Once the application is registered, configure its permissions in the Azure portal.
3. OneDrive Setup: Ensure the original Excel file is already uploaded to the root of the SharePoint document library in OneDrive. The script assumes the file's existence and will not create an initial version of it. Example template: [testing300-template.xlsx](template%2Ftesting300-template.xlsx)
4. Install Node.js and npm.
5. Install the required Node.js modules using `npm install dotenv axios exceljs`.

## Environment Variables

To configure the script properly, you need to set the following environment variables. Store these variables in a `.env` file located in the root directory of your project:

**Description of Variables:**

- `AZURE_TENANT_ID`: Your Azure AD tenant ID.
- `AZURE_CLIENT_ID`: Your Azure AD application's client ID.
- `AZURE_CLIENT_SECRET`: Your Azure AD application's client secret.
- `AZURE_SHAREPOINT_SITE_ID`: ID of the SharePoint site where the file is located.
- `AZURE_FILE_NAME`: Name of the file to be downloaded and uploaded on OneDrive.

## How to Run

1. First, create a `.env` file in the root directory of your project and set the environment variables mentioned above.
2. Run the script using the command: `node src/index.js`.

## Script Functions Breakdown

1. **getAccessToken()**: Authenticates with the Microsoft Graph API and retrieves the access token.
2. **downloadFileFromOneDrive(accessToken)**: Downloads the specified Excel file from OneDrive.
3. **uploadFileToOneDrive(accessToken, fileData)**: Uploads the updated Excel file back to OneDrive. If the file is locked, it will delete the original and create a new one.
4. **findRowByID(worksheet, candidateID)**: Finds a row in the worksheet by its ID.
5. **getFieldByName(name, candidate)**: Gets a field's value by its name from a candidate object.
6. **setCellValueAndColor(row, cellNumber, value)**: Updates a cell's value and background color.
7. **update_excel(items)**: Main function for reading, updating, and saving the Excel sheet.
8. **getDataFromAPI()**: Fetches data from an external API, transforms it, and returns it in the required format.

## Output
Example script output:
```bash
λ node src\index.js
OpenDrive data fetched!
..updating row
..updating row
..updating row
..updating row
..updating row
Current file is locked. Deleting and creating a new one.
Successfully saved to OneDrive
```

## Result
![yrYa0pccZT.jpg](imgs%2FyrYa0pccZT.jpg)

---

## Configuring Azure API Permissions:

For the Azure application, ensure that you've set it up with the "Application Permissions" mode and granted the following permission:
```bash
Sites.Selected
```

The primary advantage of the `Site.Selected` permission is that it provides granular access control. Unlike broader permissions that might grant an application access to all SharePoint sites, `Site.Selected` allows developers to restrict an application's access to only specific sites or document libraries. This ensures that the application can only access the data it genuinely needs, adhering to the principle of least privilege.

### Sharepoint permissions

It's crucial to remember that the specific SharePoint site containing the OneDrive document library must be configured with adequate "write" permissions for our `application_id`. This configuration is necessary for the script to access and modify the files.

A SharePoint administrator can grant these permissions using the endpoint:
```
${GRAPH_BASE_URL}/sites/${siteId}/permissions
```

## What happens if the file is locked?
If the file on OneDrive is locked (e.g., it's being edited by someone else), the function will first attempt to delete the locked file and then re-upload the updated file. This approach ensures that the file on OneDrive always contains the latest data, even if it means removing a currently locked file.

However, it's essential to understand the implications of this: deleting the locked file might disrupt users who are actively working on it. Ensure that this approach aligns with the requirements and usage patterns of your application.

## Fetching Data from External APIs

The script is designed to be versatile and can easily fetch data from external APIs to update the Excel sheet. One of the key strengths of the script is its ability to integrate and pull data from different sources, making it adaptable to various use-cases.

### Example: `getDataFromAPI()` Function

The `getDataFromAPI()` function serves as an excellent example of this capability. Here's a brief overview:

```javascript
async function getDataFromAPI() {
    try {
        const response = await axios.get('https://jsonplaceholder.typicode.com/users');
        return response.data.map(user => ({
            id: user.id,
            name: user.name,
            email: user.email,
            phone: user.phone,
            fields: [
                { name: 'Position Applied For', values: [{ text: user.id % 2 === 0 ? 'Web Developer' : 'System Analyst' }] },
                { name: 'Company', values: [{ text: user.company.name }] },
                { name: 'City', values: [{ text: user.address.city }] }
            ]
        }));
    } catch (error) {
        console.error('Error fetching data from API:', error);
        return [];
    }
}
```
