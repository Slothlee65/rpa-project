# INSERT ANSWER FOR QUESTION 1 HERE
while True:    
    directory = single_input("Please enter the directory path where your files are located.")
    directory = directory.strip('"')
    if directory.endswith("\\"):
        pass
    else:
        directory = directory + '\\'
    try:
        filenames = os.listdir(directory)
        if 'ProductData.xlsx' in filenames and 'MasterInvoice.pdf' in filenames and 'Order1.xlsx' in filenames and 'Order2.xlsx' in filenames and 'Order3.xlsx' in filenames and 'Order4.xlsx' in filenames:
            break
        else:
            show_message("Error", "The necessary files were not found in the directory you entered. Please try again.")
    except:
        show_message("Error", "The directory you entered is not valid. Please try again.")
directory

# INSERT ANSWER FOR QUESTION 2 HERE
df1 = pd.read_excel(directory + 'ProductData.xlsx', sheet_name = 'Descriptions')
df2 = pd.read_excel(directory + 'ProductData.xlsx', sheet_name = 'Prices')

product_descriptions = {}
product_prices = {}
for index,row in df1.iterrows():
    product_descriptions[row['Unit number']]=row['Description']

for index, row in df2.iterrows():
    product_prices[row['Unit number']]= row['Price per unit']

print(product_descriptions)
print(product_prices)

# INSERT ANSWER FOR QUESTION 3 HERE
def get_order_data(invoice_num,filename,directory):
    template = openpyxl.load_workbook(directory + filename)
        
    selected_sheet = template['OrderForm']

    order_dictionary = {}
    order_dictionary['InvoiceNum'] = str(invoice_num)
    datevalue = selected_sheet['C9'].value
    OrderDate = f"{datevalue.month:02d}/{datevalue.day:02d}/{datevalue.year:04d}"
    order_dictionary['Date'] = OrderDate
    total = 0
    order_dictionary['CustomerAddress1'] = str(selected_sheet['C16'].value)
    CustomerName = selected_sheet['C14'].value
    CustomerEmail = selected_sheet['C13'].value
    order_dictionary['CustomerName'] = str(CustomerName)
    order_dictionary['CustomerCompany'] = str(selected_sheet['C15'].value)
    order_dictionary['CustomerAddress2'] = f"{selected_sheet['C17'].value}, {selected_sheet['C18'].value} {str(selected_sheet['C19'].value)}"
    
    if selected_sheet['B24'].value is not None:
        UnitNumber0 = selected_sheet['B24'].value
        quantity0 = selected_sheet['C24'].value
        price0 = product_prices[UnitNumber0]
        amount0 = quantity0*price0
        description0 = product_descriptions[UnitNumber0]
        order_dictionary['Description0'] = str(description0)
        order_dictionary['Quantity0'] = str(quantity0)
        order_dictionary['Price0'] = f"${price0:,.2f}"
        order_dictionary['Amount0'] = f"${amount0:,.2f}"
        total = total+amount0

    if selected_sheet['B25'].value is not None:
        UnitNumber1 = selected_sheet['B25'].value
        quantity1 = selected_sheet['C25'].value
        price1 = product_prices[UnitNumber1]
        amount1 = quantity1*price1
        description1 = product_descriptions[UnitNumber1]
        order_dictionary['Description1'] = str(description1)
        order_dictionary['Quantity1'] = str(quantity1)
        order_dictionary['Price1'] = f"${price1:,.2f}"
        order_dictionary['Amount1'] = f"${amount1:,.2f}"
        total = total+amount1

    if selected_sheet['B26'].value is not None:
        UnitNumber2 = selected_sheet['B26'].value
        quantity2 = selected_sheet['C26'].value
        price2 = product_prices[UnitNumber2]
        amount2 = quantity2*price2
        description2 = product_descriptions[UnitNumber2]
        order_dictionary['Description2'] = str(description2)
        order_dictionary['Quantity2'] = str(quantity2)
        order_dictionary['Price2'] = f"${price2:,.2f}"
        order_dictionary['Amount2'] = f"${amount2:,.2f}"
        total = total+amount2

    order_dictionary['Total'] = f"${total:,.2f}"

    return order_dictionary, CustomerName, CustomerEmail, OrderDate

order_data, name, email, date = get_order_data(8515, 'Order1.xlsx',directory)
print(order_data)
print(email)
print(name)
print(date)

# INSERT ANSWER FOR QUESTION 4 HERE

def create_invoice(invoice_num,order_data,directory):
    create_pdf(directory+ 'MasterInvoice.pdf', directory + f'Invoice {invoice_num}.pdf', order_data)

create_invoice(8515,order_data,directory)

# INSERT ANSWER FOR QUESTION 5 HERE
gmail_password_path = select_file('Please select the gmail password file')
gmail_password = open(gmail_password_path).read()

def send_email(invoice_num,name,email,date,directory):

    gmail_user = "slothlee65@gmail.com"

    recipients = [email]

    subject = f'Your Order from PythonTech Co. - Invoice {invoice_num}'

    body = f"""Dear {name},
        
    Attached is your invoice dated {date}.

    Thank you for your business.
        
    PythonTech Co.
    10 Python Drive
    Pleasant Grove, UT 84062
    Phone: (801) 555-2372
    Email: sales@pythontechco.com"""

    # Initialize the email message

    message = EmailMessage()
    message['From'] = gmail_user
    message['To'] = recipients
    message['Subject'] = subject
    message.set_content(body)

    # Add attachment

    path_to_file = f"{directory}Invoice {invoice_num}.pdf"
    #path_to_file = os.getcwd()+'/Invoice Template.xlsx' # MAC
    message = add_attachment(path_to_file, message) # This function is available from the analyticskit package
            
    # Log in to the Gmail server

    mail_server = smtplib.SMTP_SSL('smtp.gmail.com')
    mail_server.login(gmail_user, gmail_password)

    # Send the email

    mail_server.send_message(message)
    mail_server.quit()

send_email(8515,name,email,date,directory)

# INSERT ANSWER FOR QUESTION 6 HERE
filenames = os.listdir(directory)
invoice_num = 8515
for file in filenames:
    if 'Order' in file and 'xlsx' in file:
        order_data,name,email,date = get_order_data(invoice_num,file,directory)
        create_invoice(invoice_num,order_data,directory)
        send_email(invoice_num,name,email,date,directory)
        invoice_num = invoice_num+1

