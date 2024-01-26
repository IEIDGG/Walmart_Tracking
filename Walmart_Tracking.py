import imaplib
import email
from bs4 import BeautifulSoup
import csv
import datetime

def extract_details(email_html):
    try:
        soup = BeautifulSoup(email_html, 'html.parser')

        links = soup.find_all('a', href=True, string=True)
        try:
            tracking_numbers = [link.string for link in links if
                                'tracking number' in link.find_previous('div').text.lower()]
        except AttributeError:
            print('tracking number error 1')
            tracking_numbers = []
        if not tracking_numbers:
            try:
                tracking_numbers = [link.string for link in links if
                                    'tracking number' in link.find_previous('span').text.lower()]
            except AttributeError:
                print('tracking number error 2')
                tracking_numbers = []

        order_date_div = soup.find('div', string=lambda text: text and 'order date' in text.lower())
        order_date = order_date_div.get_text(strip=True).split(': ')[1] if order_date_div else ''

        try:
            order_number_a = soup.find('div', string=lambda text: text and 'order number' in text.lower())
            order_number = order_number_a.get_text(strip=True).split(': ')[1] if order_number_a else ''
        except AttributeError:
            print('order number error 1')
            order_number = []
        if not order_number:
            try:
                order_number_a = soup.find('a', href=True, string=True, target=True)
                order_number = order_number_a.text
            except AttributeError:
                print('order number error 2')
                order_number = []

        try:
            address_span = soup.find('span', style='font-family:helvetica;font-size:16px!important;font-weight:400!important;color:rgb(46,47,50)')
            address_text = address_span.text
        except AttributeError:
            print('address error 1')
            address_text = ''
        if not address_text:
            try:
                address_div = soup.find('div',
                                        style='color:#6d6e71;font-family:helvetica;font-size:16px;line-height:1.38;text-align:left;text-decoration:none;')
                address_text = address_div.get_text(strip=True) if address_div else ''
            except AttributeError:
                print('address error 2, trying 3')
                address_div = soup.find_all('p', style='margin:0')
                if len(address_div) >= 3:
                    address_text = address_div[2].get_text()
                    print(address_text)
                else:
                    raise ValueError("There are not enough <p> elements in the HTML.")

        return tracking_numbers, order_number, order_date, address_text
    except Exception as e:
        print(f"Error in extract_details: {e}")
        return [], '', '', ''

def read_credentials(file_path):
    try:
        with open(file_path, 'r') as file:
            lines = file.readlines()
            if len(lines) >= 2:
                user = lines[0].split('=')[1].strip()
                password = lines[1].split('=')[1].strip()
                return user, password
            else:
                raise ValueError("The file should contain at least two lines (email and password).")
    except FileNotFoundError:
        print(f"Error: File '{file_path}' not found.")
        raise
    except Exception as e:
        print(f"Error in read_credentials: {e}")
        raise

try:
    imap_url = 'imap.gmail.com'
    file_path = 'credentials.txt'
    user, password = read_credentials(file_path)

    mail = imaplib.IMAP4_SSL(imap_url)
    mail.login(user, password)
    mail.select("inbox")

    start_date = datetime.date(2024, 1, 1)  # Year, Month, Day
    formatted_start_date = start_date.strftime("%d-%b-%Y")

    status, messages = mail.search(None, f'(FROM "help@walmart.com" SUBJECT "Shipped" SINCE "{formatted_start_date}")')
    messages = messages[0].split()

    with open('OrderDetails.csv', mode='w', newline='', encoding='utf-8') as file:
        writer = csv.writer(file)

        writer.writerow(['Tracking Number', 'Order Number', 'Order Date', 'Address'])

        for mail_id in messages:
            status, data = mail.fetch(mail_id, '(RFC822)')
            for response_part in data:
                if isinstance(response_part, tuple):
                    message = email.message_from_bytes(response_part[1])

                    if message.is_multipart():
                        for part in message.walk():
                            if part.get_content_type() == "text/html":
                                try:
                                    html_content = part.get_payload(decode=True).decode('utf-8')
                                except UnicodeDecodeError:
                                    html_content = part.get_payload(decode=True).decode('utf-8', errors='replace')
                                tracking_numbers, order_number, order_date, address_text = extract_details(html_content)
                                for number in tracking_numbers:
                                    writer.writerow([number, order_number, order_date, address_text])
                    else:
                        if message.get_content_type() == "text/html":
                            try:
                                html_content = message.get_payload(decode=True).decode('utf-8')
                            except UnicodeDecodeError:
                                html_content = message.get_payload(decode=True).decode('utf-8', errors='replace')
                            tracking_numbers, order_number, order_date, address_text = extract_details(html_content)
                            for number in tracking_numbers:
                                writer.writerow([number, order_number, order_date, address_text])

    mail.close()
    mail.logout()
except Exception as e:
    print(f"Unexpected error: {e}")

input("Press Enter to exit.")
