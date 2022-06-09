__author__ = 'mmason'
#Version 1.0
import os
import datetime
import time
import calendar
import shutil
import re
from unicodedata import category
import xml.etree.ElementTree as ET

# date becomes the current date and is then placed in MM-DD-YY format
date = time.strftime("%x").replace("/", "-")
now = datetime.datetime.now()
last_month = now.month - 1 if now.month > 1 else 12
last_months_year = now.year - 1 if now.month == 12 else now.year
month_folder = date[:2] + "-20" + date[-2:]
year = date[-2:]
queries_xml = ET.parse('queries_list.xml') #Parse the xml file into an elementtree
root = queries_xml.getroot()
query_list = []
known_query_names = []


###############################
test = True 
###############################
class Query(object):
    name = ""
    category = ""

    def __init__(self, name, category):
        self.name = name
        self.category = category

    def __str__(self):
        return self.name + " [" + self.category + "]"


class MailGroup(object):
    name = ""
    recipients = ""
    attachments = []

    def __init__(self, recipients):
        self.attachments = []
        self.recipients = recipients


def create_query_object(name, category):
    query = Query(name, category)
    return query
        

def rename(name, new_name, attach_list, i=2):
    this_name = os.path.realpath(name)
    this_new_name = os.path.realpath(new_name)
    this_attach_list = attach_list
    num = i
    try:
        os.rename(this_name, this_new_name)
        this_attach_list.append(this_new_name)
    except WindowsError:
        try:
            final_name = this_new_name[:-4] + " (" + str(num) + ")" + this_new_name[-4:]
            os.rename(this_name, final_name)
            this_attach_list.append(final_name)
        except WindowsError:
            rename(this_name, this_new_name, this_attach_list, num + 1)


def move(name, to_directory):
    move_name = name
    move_directory = to_directory
    try:
        shutil.move(move_name, move_directory)
    except shutil.Error:
        print ("Already a file with the name:" + name + "at location.")


def mailer(text, subject, recipient, cc, attachments):
    import win32com.client as win32

    list(attachments)
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = recipient
    mail.cc = cc
    mail.Subject = "PHI " + subject
    mail.HtmlBody = text
    for each in attachments:
        mail.Attachments.Add(Source=each)
    #TODO: add protection for too many email windows open
    mail.Display()


def do_query(name, new_name, destination, attach_list, i=2):
    this_name = name
    this_new_name = new_name
    this_destination = destination
    this_attach_list = attach_list
    num = i
    if num == 2:
        move(this_name, this_destination)
        rename(destination + "/" + this_name, destination + "/" + this_new_name, this_attach_list)


def make_recipients(*args):
    addresses = ""
    for i in args:
        addresses = addresses + i + ";"
    return addresses


def trim_query(query):
    trimmed = ""
    x = query.split('_')
    
    #recreate the query name
    for each in x[:-1]:
        trimmed += each + "_"
    
    #add the XX for aid years in query
    if "-" in x[-1]:
        trimmed += "XX"
    #if query does not have an aid year in its name trim leftover "_"
    else:
        trimmed = trimmed[:-1]

    return trimmed


def query_in_list(query):
    #check if trimmed is in list of queries
    if trim_query(query) in known_query_names:
        return True
    else:
        return False


def do_queries():
    global aid_year
    year = date[:2]
    aid_year = "20" + str(int(year) - 1) + "-20" + year
    for query_name in os.listdir("."):
        if query_name.startswith("UUFA_IL_REPEAT_COURSES"):
            year = str(int(re.search(r'\d+', query_name).group()))
            aid_year = "20" + str(int(year) - 1) + "-20" + year
            break
    if test:
        directory = os.path.realpath(os.path.join('C:\Testing Bob/Daily', aid_year, month_folder))
        royall_directory = os.path.realpath('C:\Testing Bob/Royall')
        pell_directory = os.path.realpath(os.path.join('C:\Testing Bob/QUERIES\Pell Repackaging', aid_year))
        disb_directory = os.path.realpath('C:\Testing Bob\QUERIES\Disbursement\Pre-Disbursement Queries')
        refund_directory = os.path.realpath(os.path.join('C:\Testing Bob/QUERIES\Refund Credit Holds', month_folder))
    else:
        directory = os.path.realpath(os.path.join('O:/Systems/QUERIES/Daily', aid_year, month_folder))
        royall_directory = os.path.realpath('O:/Systems/Royall')
        pell_directory = os.path.realpath(os.path.join('O:\Systems\QUERIES\Pell Repackaging', aid_year))
        disb_directory = os.path.realpath("O:\Systems\QUERIES\Disbursement\Pre-Disbursement Queries")
        refund_directory = os.path.realpath(os.path.join('O:\Systems\QUERIES\Refund Credit Holds', month_folder))

    # the list 'my_path' should be populated with the FOLDER variables above.
    if not os.path.isdir(directory):
        os.makedirs(directory)
    if not os.path.isdir(royall_directory):
        os.makedirs(royall_directory)
    if not os.path.isdir(pell_directory):
        os.makedirs(pell_directory)
    if not os.path.isdir(disb_directory ):
        os.makedirs(disb_directory)
    if not os.path.isdir(refund_directory):
        os.makedirs(refund_directory)


def main():
    #Create list of queries from xml file
    for query in root.findall('query'):
        query_list.append(create_query_object(query.find('name').text, query.find('category').text))
        known_query_names.append(query.find('name').text)
    
    #check files in root directory for match in query_list, if found process file.
    for file_name in os.listdir("."):
        if file_name.startswith("UUFA"):
            if query_in_list(file_name):
                do_query(file_name)


if __name__ == "__main__":
    main()


input("So Long, and Thanks for All the Fish.\nPRESS ENTER TO CLOSE.")