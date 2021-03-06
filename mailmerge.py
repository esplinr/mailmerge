#! /usr/bin/python3
# -*- coding: utf-8 -*-
''' mailmerge.py

Accept a CSV and an email template
The first line of the CSV are fields to substitute into the email template.

TODO: The program is currently hard-coded to relay through gmail instead of sending direct. At some point it should take that information as arguments.
'''


import sys
import email
import time

def prepare_email_message(template, value_dict):
    '''
    Take the email template and build an Email Message object.
    Substitute the value_dict for keys in the message.
    '''
    new_template="" # Needs to be a string to build the message
    for line in template:
        newline = line
        for key, value in value_dict.items():
            newline=newline.replace("#{0}#".format(key), value)
        new_template=new_template + newline
    prepared_message = email.message_from_string(new_template)
    return prepared_message


def send_emails(template, values_dict_array):
    '''
    Send a personalized email for each entry in values_dict_array.
    Substitute into the template the values in the dict

    template: an array of strings containing the lines in the email.
      First lines are (in order):
        From
        To
        Subject
      Remaining lines make up the body.
      At any time a value of #keyname# will be substituted with the matching
      value from the value_dict.

    value_dict_array: Each row contains values that match keys used in the email.
    '''
    from smtplib import SMTP, SMTP_SSL
    import smtplib
    send_count=0

    smtp_host=""
    smtp_port=""
    smtp_login=""
    smtp_pass=""
    TLS = False
    smtp = None # Can't use SMTP "with" statement, or loop will break on disconnect

    def smtp_connect():
        if TLS: # TLS on port 587
            smtp=SMTP(host=smtp_host, port=smtp_port)
            smtp.starttls()
            smtp.ehlo()
        else: # SSL on port 465
            smtp=SMTP_SSL(host=smtp_host, port=smtp_port)
        smtp.login(smtp_login, smtp_pass)
        print("connected to SMTP")
        return smtp

    try:
        smtp = smtp_connect() 
        for values_dict in values_dict_array:
            email_message = prepare_email_message(template, values_dict)
            to_addr = [t[1] for t in email_message.items() if t[0]=='To'][0]
            print("Sending to {0}".format(to_addr))
            try:
                email_message.set_charset('utf-8')
                smtp.send_message(email_message)
            except smtplib.SMTPServerDisconnected:
                print("Reconnecting to SMTP Server")
                smtp = smtp_connect()
                smtp.send_message(email_message)
            except smtplib.SMTPSenderRefused:
                print("Pausing for rate-limit")
                time.sleep(180)
                print("Reconnecting to SMTP Server")
                smtp = smtp_connect()
                smtp.send_message(email_message)
            send_count+=1
    finally:
        if smtp: smtp.quit()
    print("Emails sent: {0}".format(send_count))


def main(arg_list):
    # Setup argument parsing
    import argparse
    parser = argparse.ArgumentParser(description="Send a personalized email to each person in a CSV.")
    parser.add_argument("csv_filename",
                        help="Filename of the CSV file defining the key / value "
                             "pairs. One row per email. First row lists the "
                             "keys, each additional row lists the values.")
    parser.add_argument("email_template_filename",
                        help="Text file containing the email. #keyname# will "
                             "be substituted for values in the CSV. Required "
                             "lines: 'From:', 'To:', and 'Subject:'")
    args=parser.parse_args(arg_list)
  
    # Read the email template
    with open(args.email_template_filename, encoding='utf-8-sig') \
            as email_template_file:
        email_template = email_template_file.readlines()
  
    # Read the CSV
    import csv
    with open(args.csv_filename) as csvfile:
        dialect = csv.Sniffer().sniff(csvfile.read(4098))
        csvfile.seek(0)
        email_dict_array = csv.DictReader(csvfile, dialect=dialect)
        # Send the emails, substituting the variables for the placeholders
        send_emails(email_template, email_dict_array)
 
    return(0)


if __name__ == '__main__':
    sys.exit(main(sys.argv[1:])) #Arg list needs to drop program name


# vim: tabstop=8 expandtab shiftwidth=4 softtabstop=4
