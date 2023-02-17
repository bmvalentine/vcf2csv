"""
This script will parse a vcf file and print the names and addresses of the contacts found in the vcf.

The data is exported in a csv format.

The purpose is to create a csv that can be imported into a worksheet which can them be mail merged into sheets of address
labels.
"""

import sys
import re

def print_hi(name):
    # Use a breakpoint in the code line below to debug your script.
    print(f'Hi, {name}')  # Press ⌘F8 to toggle the breakpoint.

def print_file(name):
    file = open(name, "rt")
    print("\"First Name\",\"Last Name\",\"Addr Type\",\"Address\",\"City\",\"State\",\"Zip\"")

    while True:
        line = file.readline()
        if line == "" :
            break

        if "VERSION:3.0" in line:
            # Read the first and last names
            #
            contact_data = file.readline().split(";")
            last_name = contact_data[1].split(":")[1].strip()
            first_name = contact_data[2].strip()
            #print(first_name, last_name)
        elif re.search("^TEL", line) is not None :
            # Read the telephone number
            #
            tel_data = line.split(";")
            phone1 = tel_data[2].split(":")[1]
            #print("phone number: " + phone1[:100])
        elif re.search("^ADR", line) is not None:
            if "CHARSET=UTF-8;HOME" in line:
                addr_type = "HOME"
            else:
                addr_type = "WORK"
            addr_data = line.split(";")
            zip = addr_data[-2].strip()
            state = addr_data[-3].strip()
            city = addr_data[-4].strip()
            addr = addr_data[-5].strip().replace("\\n", " ")
            addr2 = addr_data[-6].strip().replace("\\n", " ")
            if addr2 != "":
                addr = addr + " " + addr2

            sys.stdout.write(f"\"{first_name}\",\"{last_name}\",\"{addr_type}\",\"{addr}\",\"{city}\",\"{state}\",\"{zip}\"\n")


        #sys.stdout.write( line[:100] )

    file.close()
    return

# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    print_hi('Preston')
    #print_file("Brandy Contact Backup 2022-11-27.vcf")

# See PyCharm help at https://www.jetbrains.com/help/pycharm/
