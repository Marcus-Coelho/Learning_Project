# Learning_Project

### Documentation

This script automates the process of generating sales reports for different stores, calculating performance metrics, and sending these reports via email to the store managers and management. Below are the detailed steps and functionalities:

1. **Import Libraries**: Necessary libraries are imported, including pandas for data manipulation, pathlib for handling file paths, locale for setting regional settings, smtplib for sending emails, and others.

2. **Set Locale**: The locale is set to Brazilian Portuguese for formatting numbers in the Brazilian currency format.

3. **File Reading Functions**: Functions `read_csv_with_encodings` and `read_excel_with_encodings` are defined to read CSV and Excel files with different encodings.

4. **File Paths**: Paths to the data files (emails, stores, and sales) are defined.

5. **Create DataFrames**: DataFrames are created by reading the respective files.

6. **Create Table for Each Store**: A dictionary of DataFrames is created, one for each store.

7. **Create Backup Directories**: Backup directories for each store are created if they do not already exist.

8. **Update DataFrames**: The sales DataFrames are updated with additional information from the stores and emails DataFrames.

9. **Define Goals**: Performance goals for daily and annual sales, product quantities, and average ticket values are defined.

10. **SMTP Server Configuration**: The SMTP server is configured for sending emails.

11. **Process Each Store**: For each store, the script:
    - Calculates various metrics (e.g., daily and annual sales, product quantities, average tickets).
    - Generates an HTML email body with these metrics.
    - Sends the email to the respective store manager.
    - Accumulates the email content for later use.

12. **Combine HTML for Management Email**: Combines the HTML content from all store emails into one HTML document, ensuring the DOCTYPE and other header tags appear only once.

13. **Send Email to Management with All Store Data**: Sends the combined HTML document to the management.

14. **Create and Send Rankings Email**: 
    - Creates a DataFrame with the collected data.
    - Sorts the DataFrame to generate rankings for annual and daily sales.
    - Converts the rankings to HTML and sends them in an email to the management.

This script ensures that all sales data is processed, formatted, and sent efficiently to the relevant stakeholders.
