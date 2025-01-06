# Web Automation for Appliance Data Upload

This script automates the process of logging into a web-based system to upload appliance-related data (such as manuals, images, and warranty guides) from an Excel file. The automation is done using Selenium WebDriver and Python, and it interacts with various web elements to input data and submit forms for each appliance.

## Installation

To run this script, you need the following steps:

```bash

git clone 
cd /path/to/your/project
python -m venv venv
.\venv\Scripts\activate
pip install -r requirements.txt
python script_name.py
```

## Functionality

The script performs the following steps:

**Login to the Website:**

Logs into the MoveEasy account management website with the provided credentials.
Uses the Selenium WebDriver to locate the login fields, fill them in, and submit the form.
Navigate to the Missing Manual Page:

After logging in, the script navigates to the "Missing Appliance Data List" page and the "Missing Manual View" page for uploading data.

**Read Data from Excel:**

Reads appliance data from an Excel file using Pandas. Each row contains appliance details like model number, brand, name, image URL, user manual, energy guide, and warranty guide.

**Search and Upload Data:**

For each appliance entry, it searches for the appliance by its model number in the application and uploads:

- Images:

    Downloads the image from the provided URL and uploads it to the corresponding field.

- User Manuals:

    Downloads the user manual PDF from the provided URL and uploads it.

- Energy Guide and Warranty Guide:

    Similar to the user manual, these PDFs are downloaded and uploaded.
- Handle Pagination:

    If the model is not found on the current page, the script automatically navigates through the pagination until it finds the model or attempts another strategy for finding the model.

- Submit Data:

    Once the appliance data is filled in, the script clicks the "Data Update" button to submit the form.
The page refreshes, and the status is updated for the appliance.

## Error Handling

- If the script encounters any issues while finding the appliance on the website (such as the model not existing on the page), it will attempt to navigate to the next page.
- If there are any issues downloading the image or PDF files, the script will skip that step and move forward.

## Output

- The script saves downloaded files (manuals, images, PDFs) in the ajmadison_com folder.
- It outputs logs to the console, detailing the current appliance being processed, any errors, and successful uploads.

## Important Notes

- Make sure that the Excel file provided has the necessary columns (model number, name, brand, etc.).
- The paths and file handling are specific to Windows. Adjust paths if using Linux or macOS.
- Be careful with credentials in the script, as hardcoding sensitive information (e.g., username, password) is not secure for production use.

## Product Data Table

This table represents the product data including model name, brand, category, image URL, user manual, energy guide, and warranty guide.

| **Model Name**   | **Name** | **Brand** | **Category** | **Image URL** | **User Manual** | **Energy Guide** | **Warranty Guide** |
|------------------|----------|-----------|--------------|---------------|-----------------|------------------|---------------------|
| MDB4409PAWO      | XYZ      | VIKING    | Range        | [Image Link](http://...)  | [User Manual Link](http://...) | [Energy Guide Link](http://...) | [Warranty Guide Link](http://...) |

## Future Improvements

- Secure handling of login credentials (e.g., using environment variables or a secure vault).
- Add logging for better error tracking and debugging.
