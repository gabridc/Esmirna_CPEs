from typing import Sized
from openpyxl import Workbook, load_workbook
import sys
from colorama import Fore, Back, Style
import re

def main():

    if(len(sys.argv) != 3):
        print(Fore.RED + "Command Error: py main.py <path_excel_file.xls> <path_cpe_list_file.txt>")
        print(Style.RESET_ALL)
        return 1
    
    rpmFile = open(sys.argv[1])
    cpeFile = open(sys.argv[2])

    rpms = rpmFile.readlines()
    cpes = cpeFile.readlines()

    for rpm in rpms:
        #print(rpm)
        rpm = rpm.split(";")[0].replace("++", "")
        
        patterns = []
        patterns.append("^cpe:2\.3:[a]?[o]?[h]?:"  + rpm + ":" + rpm + ":[a-zA-Z0-9\.-_*]+:*:*:*:*:*:*:*")
        patterns.append("^cpe:2\.3:[a]?[o]?[h]?:[a-zA-z0-9_\.-]+:" + rpm + ":[a-zA-Z0-9\.-_*]+:*:*:*:*:*:*:*")

        if  len(rpm.split("-")) == 3:
           patterns.append("^cpe:2\.3:[a]?[o]?[h]?:[a-zA-z0-9_\.-]+:" + rpm.split("-")[0] + "-" + rpm.split("-")[1]  + ":[a-zA-Z0-9\.-_*]+:*:*:*:*:*:*:*")
           patterns.append("^cpe:2\.3:[a]?[o]?[h]?:[a-zA-z0-9_\.-]+:" + rpm.split("-")[0].rstrip(rpm.split("-")[0][-1]) + ":[a-zA-Z0-9\.-_*]+:*:*:*:*:*:*:*")
        elif len(rpm.split("-")) == 2:
            patterns.append("^cpe:2\.3:[a]?[o]?[h]?:[a-zA-z0-9_\.-]+:" + rpm.split("-")[0] + ":[a-zA-Z0-9\.-_*]+:*:*:*:*:*:*:*")
            patterns.append("^cpe:2\.3:[a]?[o]?[h]?:[a-zA-z0-9_\.-]+:" + rpm.split("-")[0].rstrip(rpm.split("-")[0][-1]) + ":[a-zA-Z0-9\.-_*]+:*:*:*:*:*:*:*")
            

        for pattern in patterns:
            r = re.compile(pattern)
            newlist = list(filter(r.match, cpes)) # Read Note
            if  len(newlist) > 0:
                print(rpm + " " + " [" + str(len(newlist))  + "] " + newlist[0].replace("\n", "") + " Pattern: " +  pattern)
                break
            else:
                print(rpm)
             


    #cpes = ["ssslskype", "hh"]
    #print(type(cpes))

    
    newlist = list(filter(r.match, cpes)) # Read Note

    #print(newlist)



if __name__ == "__main__":
    main()