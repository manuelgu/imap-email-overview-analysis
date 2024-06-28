import email
import imaplib
from email.header import decode_header

import pandas as pd
import pytz
import yaml

with open('config.yaml', 'r') as file:
    config = yaml.safe_load(file)

# Account credentials
username = config["username"]
password = config["password"]

# Create an IMAP4 class with SSL
mail = imaplib.IMAP4_SSL(config["host"])

# Authenticate
mail.login(username, password)

# Select the mailbox you want to use (in this case, the inbox)
mail.select(config["inbox_name"])

# Search for all emails in the inbox
status, messages = mail.search(None, "ALL")

# Convert messages to a list of email IDs
email_ids = messages[0].split()

# Initialize a dictionary to store email data
email_data = {}

timezone = pytz.timezone('Europe/Stockholm')

# Fetch the email message by ID
for email_id in email_ids:
    # Fetch the email body (RFC822) and flags for the given ID without marking it as read
    status, msg_data = mail.fetch(email_id, "(BODY.PEEK[] FLAGS)")

    for response_part in msg_data:
        if isinstance(response_part, tuple):
            msg = email.message_from_bytes(response_part[1])

            # Decode the email subject
            subject, encoding = decode_header(msg["Subject"])[0]
            if isinstance(subject, bytes):
                # If it's a bytes type, decode to str
                subject = subject.decode(encoding if encoding else "utf-8")

            # Email sender
            from_ = msg.get("From")
            name, email_address = email.utils.parseaddr(from_)

            # Decode the sender's name if it is encoded
            decoded_name = decode_header(name)
            name_parts = []
            for part, enc in decoded_name:
                if isinstance(part, bytes):
                    name_parts.append(part.decode(enc if enc else 'utf-8'))
                else:
                    name_parts.append(part)
            name = ''.join(name_parts)

            # Check if the email has been read
            flags = response_part[0]
            is_read = "\\Seen" in flags.decode()

            # Email date
            date_ = msg.get("Date")
            date_parsed = email.utils.parsedate_to_datetime(date_).astimezone(timezone)

            if email_address not in email_data:
                email_data[email_address] = {
                    "received_email_count": 0,
                    "read_email_count": 0,
                    "names": set([]),
                    "earliest_date": date_parsed,
                    "latest_date": date_parsed,
                    "last_read_email_date": None
                }

            email_data[email_address]["received_email_count"] += 1
            if is_read:
                email_data[email_address]["read_email_count"] += 1
                if (email_data[email_address]["last_read_email_date"] is None or
                        date_parsed > email_data[email_address]["last_read_email_date"]):
                    email_data[email_address]["last_read_email_date"] = date_parsed

            email_data[email_address]["earliest_date"] = min(email_data[email_address]["earliest_date"], date_parsed)
            email_data[email_address]["latest_date"] = max(email_data[email_address]["latest_date"], date_parsed)
            email_data[email_address]["names"].add(name)

# Convert dictionary to DataFrame
df = pd.DataFrame.from_dict(email_data, orient='index')
df.reset_index(inplace=True)
df.rename(columns={"index": "email_address"}, inplace=True)

# Sort by 'received_email_count'
df.sort_values(by="received_email_count", ascending=False, inplace=True)

# Convert dates to string format for CSV
df["earliest_date"] = df["earliest_date"].dt.strftime("%Y-%m-%d %H:%M:%S")
df["latest_date"] = df["latest_date"].dt.strftime("%Y-%m-%d %H:%M:%S")
df["last_read_email_date"] = df["last_read_email_date"].dt.strftime("%Y-%m-%d %H:%M:%S")

# Save DataFrame to CSV
df.to_csv("email_data.csv", index=False)

# Close the connection and logout
mail.close()
mail.logout()
