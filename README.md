Download data from tijori.com
This VBA code automates the process of collecting financial data from a website and organizing it in an Excel spreadsheet. It leverages the following functionalities:

Data Scraping:

Utilizes a WebDriver object to interact with a web browser (likely Chrome).
Navigates to a specific URL (provided in cell E2 of Sheet2).
Extracts data elements based on class names ("details-control") and tags ("tr", "td", "th").
Filters extracted data based on the presence of the "+" symbol.
Data Organization:

Writes extracted data (from "td" elements) to column A of Sheet2, skipping empty cells.
Writes extracted data (from "th" elements) to column B of Sheet2, skipping empty cells.
Copies relevant data from Sheet2 to Sheet1, handling different data formats (numeric, text with units like "Rs" or "Lakh").
Inserts rows and headers in Sheet1 to categorize the data (Balance Sheet, Profit & Loss, etc.).
Prompts the user to select specific data ranges (year, companies, quarters) for formatting in Sheet1.
Overall, this code streamlines the process of gathering and structuring financial data from a website.

Note:

The code relies on external libraries (WebDriver) and might require additional setup.
Modifying the code to work with different website structures might be necessary.
Refer to the linked LinkedIn article (https://www.linkedin.com/posts/sayandevnandy_an-afternoon-well-spent-on-discussing-the-activity-7143141568026189824-KiKX) for a more detailed explanation and potential adaptations.
