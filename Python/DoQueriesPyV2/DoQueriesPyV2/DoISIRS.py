import os
import csv
import datetime


file_list = []
todays_date = datetime.date.today()
isir_directory = "H:\FA\T4WAN\data"

class ISIR(object):
    name = ""
    count = 0;
    date = datetime.date

    def __init__(self, date, name, count):
        self.date = date
        self.name = name
        self.count = count

    def __str__(self):
        return str(self.date) + "," + str(self.name) + "," + str(self.count)
        

def create_isir(date, name, count):
    isir = ISIR(date, name, count)
    return isir


def count_isir(fname):
    name = os.path.basename(fname)
    date = datetime.date.fromtimestamp(os.path.getmtime(fname))
    num_lines = sum(1 for line in open(fname))-1
    file_list.append(create_isir(date,name,num_lines))


def main():
    for fname in os.listdir("."):
        fname_date = datetime.date.fromtimestamp((os.path.getmtime(os.path.join(".",fname))))
        #if(fname_date == todays_date):
        if fname.startswith("igsa") or fname.startswith("idsa") or fname.startswith("igco") or fname.startswith("igsg"):
            count_isir(os.path.join(".",fname))

    #TODO: create a csv file with the data to paste into the master log.
    #TODO: create/edit the ISIRS.txt file to use with ISIR Import
    for item in file_list:
        print(item)


if __name__ == "__main__":
    main()
