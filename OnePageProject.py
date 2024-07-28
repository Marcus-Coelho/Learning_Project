# %%
import pandas as pd
from pathlib import Path
import locale
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import time
import random

# %%
# Set locale to Brazil
locale.setlocale(locale.LC_ALL, "pt_BR.UTF-8")

# %%
# Functions to read files with different encodings
def read_csv_with_encodings(file_path):
    encodings = ["latin1", "iso-8859-1", "utf-8"]
    for encoding in encodings:
        try:
            df = pd.read_csv(file_path, sep=";", encoding=encoding)
            return df
        except Exception as e:
            print(f"Failed to read {file_path} with encoding {encoding}: {e}")
    raise ValueError(f"Could not read {file_path} with any of the provided encodings")

def read_excel_with_encodings(file_path):
    encodings = ["utf-8", "latin1", "iso-8859-1", "ascii", "utf-16", "utf-32", "cp1252"]
    for encoding in encodings:
        try:
            df = pd.read_excel(file_path)
            return df
        except Exception as e:
            print(f"Failed to read {file_path} with encoding {encoding}: {e}")
    raise ValueError(f"Could not read {file_path} with any of the provided encodings")

# %%
# File paths
emails_path = r"File Folder Adress"
stores_path = r"File Folder Adress"
sales_path = r"File Folder Adress"

# %%
# Create dataframes
dfEmails = read_excel_with_encodings(emails_path)
dfStores = read_csv_with_encodings(stores_path)
dfSales = read_excel_with_encodings(sales_path)

# %%
# Creating a table for each store
store_ids = sorted(dfSales["Store ID"].unique())
store_dataframes_dict = {store: dfSales[dfSales["Store ID"] == store] for store in store_ids}

# %%
# Define the base path of the directory
backup_path = Path(r"folder File where Beckup done")
backup_path.mkdir(parents=True, exist_ok=True)

for store in dfStores["Store"]:
    store_path = backup_path / store
    store_path.mkdir(parents=True, exist_ok=True)

# Updating the dictionary of store sales dataframes
for store_id in store_dataframes_dict:
    store_dataframes_dict[store_id] = store_dataframes_dict[store_id].merge(dfStores, on="Store ID", how="left")
    store_dataframes_dict[store_id] = store_dataframes_dict[store_id].merge(dfEmails, on="Store", how="left")

# %%
# Definitions of goals
daily_sales_goal = 0
annual_sales_goal = 0
daily_product_qty_goal = 0
annual_product_qty_goal = 0
daily_avg_ticket_goal = 0
annual_avg_ticket_goal = 0

# %%
# Define the sender and recipient details
sender = "your_email@gmail.com"
password = "password gmail"
# SMTP server configuration
server = smtplib.SMTP("smtp.gmail.com", 587)
server.starttls()
server.login(sender, password)

# %%
# Variable to accumulate results for the management
result_lines_management = []

# Lists to collect the data
store_name_list = []
annual_sales_list = []
daily_sales_list = []

# Initialization of ordered dataframes
annual_sales_ranking_list = {}
daily_sales_ranking_list = {}

# %%
for dataframe in store_dataframes_dict:
    max_date = store_dataframes_dict[dataframe]["Date"].max()
    store_df = store_dataframes_dict[dataframe]
    store_name = store_dataframes_dict[dataframe]["Store"].iloc[1]
    manager_name = store_dataframes_dict[dataframe]["Manager"].iloc[1]

    # Sales - sum the "Final Value" column
    annual_sales = store_df["Final Value"].sum()
    daily_sales = store_df[store_df["Date"] == max_date]["Final Value"].sum()
    daily_product_qty = store_df[store_df["Date"] == max_date]["Quantity"].sum()
    annual_product_qty = store_df["Quantity"].sum()

    # Appending data to the lists
    store_name_list.append(store_name)
    annual_sales_list.append(annual_sales)
    daily_sales_list.append(daily_sales)

    # Product diversity - count unique values in the "Product" column
    product_diversity = store_df["Product"].nunique()

    # Average ticket - count unique values in the "Sale Code" column and divide the sales by quantTickets
    ticket_count = store_df["Sale Code"].nunique()
    annual_avg_ticket = annual_sales / ticket_count
    daily_avg_ticket = daily_sales / daily_product_qty

    # Calculating status
    status_daily_sales = "✅" if daily_sales >= daily_sales_goal else "❌"
    status_annual_sales = "✅" if annual_sales >= annual_sales_goal else "❌"
    status_daily_product_qty = "✅" if daily_product_qty >= daily_product_qty_goal else "❌"
    status_annual_product_qty = "✅" if annual_product_qty >= annual_product_qty_goal else "❌"
    status_daily_avg_ticket = "✅" if daily_avg_ticket >= daily_avg_ticket_goal else "❌"
    status_annual_avg_ticket = "✅" if annual_avg_ticket >= annual_avg_ticket_goal else "❌"

    # Generating the email body in HTML
    html_stores = f"""
    <!DOCTYPE html>
    <html>
    <head>
        <style>
            table {{ width: 100%;border-collapse: collapse;}}
            th, td {{padding: 8px;text-align: left;border: 1px solid #ddd;}}
            th {{background-color: #f2f2f2;}}
            .goal-met {{color: green;}}
            .goal-not-met {{color: red;}}
        </style>
    </head>
    <body>
        <p>Hello {manager_name}, we present the sales summary for {max_date.strftime('%d/%m/%Y')} in the table below, as well as the performance of the indicators.</p>
        <p><strong>Store:</strong> {store_name}</p>
        <table>
            <tr>
                <th>Metric</th>
                <th>Value</th>
                <th>Goal</th>
                <th>Status</th>
            </tr>
            <tr>
                <td>Sales on {max_date.strftime('%d/%m/%Y')}</td>
                <td>{locale.currency(daily_sales, grouping=True)}</td>
                <td>{locale.currency(daily_sales_goal, grouping=True)}</td>
                <td style="text-align: center;">{status_daily_sales}</td>
            </tr>
            <tr>
                <td>Product quantity sold on the day</td>
                <td>{daily_product_qty}</td>
                <td>{daily_product_qty_goal}</td>
                <td style="text-align: center;">{status_daily_product_qty}</td>
            </tr>
            <tr>
                <td>Average ticket of the day</td>
                <td>{locale.currency(daily_avg_ticket, grouping=True)}</td>
                <td>{locale.currency(daily_avg_ticket_goal, grouping=True)}</td>
                <td style="text-align: center;">{status_daily_avg_ticket}</td>
            </tr>
            <tr>
                <td>Annual sales</td>
                <td>{locale.currency(annual_sales, grouping=True)}</td>
                <td>{locale.currency(annual_sales_goal, grouping=True)}</td>
                <td style="text-align: center;">{status_annual_sales}</td>
            </tr>
            <tr>
                <td>Product diversity</td>
                <td>{product_diversity}</td>
                <td>{annual_product_qty_goal}</td>
                <td style="text-align: center;">{status_annual_product_qty}</td>
            </tr>
            <tr>
                <td>Product quantity sold in the year</td>
                <td>{annual_product_qty}</td>
                <td>{annual_product_qty_goal}</td>
                <td style="text-align: center;">{status_annual_product_qty}</td>
            </tr>
            <tr>
                <td>Average ticket in the year</td>
                <td>{locale.currency(annual_avg_ticket, grouping=True)}</td>
                <td>{locale.currency(annual_avg_ticket_goal, grouping=True)}</td>
                <td style="text-align: center;">{status_annual_avg_ticket}</td>
            </tr>
        </table>
    </body>
    </html>
    """
    # Send email to each store
    recipient = store_dataframes_dict[dataframe]["E-mail"].iloc[1]
    msg = MIMEMultipart()
    msg["From"] = sender
    msg["To"] = recipient
    msg["Subject"] = f"One Page store {store_name} on {max_date.strftime('%d/%m/%Y')}"

    msg.attach(MIMEText(html_stores, "html"))
    server.sendmail(sender, recipient, msg.as_string())

    # Use time.sleep so that the Google server doesn't consider the submission to be SPANS
    time.sleep(random.randint(40, 68))

    print(f"Email sent to {store_name} ({manager_name}) - {recipient}")

    # Accumulate results for the management
    result_lines_management.append(html_stores)

# %%
# Function to remove the greeting paragraph
def remove_greeting(html_content):
    start_greeting = html_content.find("<p>Hello")
    end_greeting = html_content.find("</p>", start_greeting) + 4
    if start_greeting != -1 and end_greeting != -1:
        return html_content[:start_greeting] + html_content[end_greeting:]
    return html_content

# Combine HTML strings, ensuring the DOCTYPE and other header tags appear only once
combined_html = remove_greeting(result_lines_management[0])

for html_part in result_lines_management[1:]:
    # Remove the DOCTYPE and other header tags from the subsequent parts
    html_body = remove_greeting(html_part).split("<body>")[1].split("</body>")[0]
    combined_html += html_body

# Close the HTML document properly
combined_html += "</body></html>"

# %%
# Send email to management with all store data
recipient = "recipient_manager_email@gmail.com"
msg = MIMEMultipart()
msg["From"] = sender
msg["To"] = recipient
msg["Subject"] = f"One Page of Stores on {max_date.strftime('%d/%m/%Y')}"

msg.attach(MIMEText(combined_html, "html"))
server.sendmail(sender, recipient, msg.as_string())
time.sleep(random.randint(45, 68))

print("Email with all store data sent to management")

# %%
# Send email to management with rankings

# Creating the Ranking DataFrame after collecting all data
ranking_list = pd.DataFrame({
    "Store Name": store_name_list,
    "Annual Sales": annual_sales_list,
    "Daily Sales": daily_sales_list
})

# Sort ranking_list by "Annual Sales" in descending order
ranking_list_annual_sales = ranking_list.sort_values(by="Annual Sales", ascending=False)

# Sort ranking_list by "Daily Sales" in descending order
ranking_list_daily_sales = ranking_list.sort_values(by="Daily Sales", ascending=False)

# Format values in Brazilian currency
ranking_list_annual_sales["Annual Sales"] = ranking_list_annual_sales["Annual Sales"].apply(lambda x: locale.currency(x, grouping=True))
ranking_list_annual_sales["Daily Sales"] = ranking_list_annual_sales["Daily Sales"].apply(lambda x: locale.currency(x, grouping=True))
ranking_list_daily_sales["Daily Sales"] = ranking_list_daily_sales["Daily Sales"].apply(lambda x: locale.currency(x, grouping=True))
ranking_list_daily_sales["Annual Sales"] = ranking_list_daily_sales["Annual Sales"].apply(lambda x: locale.currency(x, grouping=True))

# Convert DataFrames to HTML
html_annual_sales = ranking_list_annual_sales.to_html(index=False, classes='annual_sales')
html_daily_sales = ranking_list_daily_sales.to_html(index=False, classes='daily_sales')

# Prepare email content with CSS style
html_content = f"""
<html>
<head>
<style>
    table {{width: 100%;border-collapse: collapse;}}
    th, td {{border: 1px solid black;padding: 8px;text-align: left;}}
    th {{background-color: #f2f2f2;}}
    .annual_sales {{border: 2px solid #4CAF50;}}
    .daily_sales {{border: 2px solid #f44336;}}
</style>
</head>
<body>
    <h3>RANKING ANNUAL SALES</h3>
    {html_annual_sales}
    <h3>RANKING DAILY SALES</h3>
    {html_daily_sales}
</body>
</html>
"""

recipient = "recipient_manager_email@gmail.com"
msg = MIMEMultipart()
msg["From"] = sender
msg["To"] = recipient
msg["Subject"] = f"One Page of Stores on {max_date.strftime('%d/%m/%Y')}"

msg.attach(MIMEText(html_content, "html"))
server.sendmail(sender, recipient, msg.as_string())

print("Email with rankings sent to management")
print("nothint")
# Close the connection
server.quit()
