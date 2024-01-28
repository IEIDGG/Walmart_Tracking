import imaplib
import email
from bs4 import BeautifulSoup
import csv
import datetime


def extract_details(email_html):
    try:
        soup = BeautifulSoup(email_html, 'html.parser')

        links = soup.find_all('a', href=True, string=True)

        def tracking_finder():
            tracking = [link.string for link in links if
                                    'tracking number' in link.find_previous('div').text.lower()]
            if not tracking:

                print('Tracking - Moving to method 2')

                tracking_element = [link.string for link in soup.findAll('a', style='font-family:helvetica', target="_blank", href=True, rel=False)]
                if 'Help Center' in tracking_element:
                    tracking_element.remove('Help Center')
                tracking = tracking_element

            print(tracking)

            return tracking

        def order_date_finder():
            order_date_div = soup.find('div', string=lambda text: text and 'order date' in text.lower())
            order_date = order_date_div.get_text(strip=True).split(': ')[1] if order_date_div else ''

            return order_date

        def order_number_finder():
            order_number_div = soup.find('div', string=lambda text: text and 'order number' in text.lower())
            order_number = order_number_div.get_text(strip=True).split(': ')[1] if order_number_div else ''

            if not order_number:
                order_number_a = soup.find('a', href=True, string=True, target=True)
                order_number = order_number_a.text

            return order_number

        def address_finder():
            address_div = soup.find('div', style='color:#6d6e71;font-family:helvetica;font-size:16px;line-height:1.38;text-align:left;text-decoration:none;')
            address = address_div.get_text(strip=True) if address_div else ''

            if not address:
                print('Address - Moving to method 2')
                address_span = soup.find('span', style='font-family:helvetica;font-size:16px!important;font-weight:400!important;color:rgb(46,47,50)')
                address = address_span.text

                if not address:
                    print('Address - Moving to method 3')
                    address_div = soup.find_all('p', style='margin:0')
                    if len(address_div) >= 3:
                        address = address_div[2].get_text()
                        print(address_text)
                    else:
                        print('Address - Method 3 failed, returning fail')
                        address = 'Failed to retrieve address'

            return address

        return tracking_finder(), order_date_finder(), order_number_finder(), address_finder()
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

    status, messages = mail.search(None,
                                   f'(FROM "alertspoiler00.93@gmail.com" SUBJECT "Shipped" SINCE "{formatted_start_date}")')
    messages = messages[0].split(b' ')

    with open('OrderDetails.csv', mode='w', newline='', encoding='utf-8') as file:
        writer = csv.writer(file)

        writer.writerow(['Tracking Number', 'Order Number', 'Order Date'])

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
except:
    print(f"Unexpected error: {e}")

input("Press Enter to exit.")
