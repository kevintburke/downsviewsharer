import requests
import pandas as pd
import tkinter as tk
from tkinter import messagebox, filedialog, ttk
import xml.etree.ElementTree as ET
import datetime
import time
import re
from PyZ3950 import zoom
import os

###PyZ3950 is no longer maintained(?) and includes errors/outdated lines that break functionality. Required updates to restore basic function included updating regex syntax in zoom.py (switching (?i)... for (?i:...)), turning off strict parsing in zmarc.py, and adding a step to zmarc.py to skip all non-numeric fields

###REFORMAT TO RETURN ORIGINAL SPREADSHEET WITH SUCCESS OR FAILURE AND REASON IN NEW COLUMNS???

global BASEURL
BASEURL = "https://api-ca.hosted.exlibrisgroup.com/almaws/v1/"

class DownsviewSharer:
    def __init__(self, root):
        self.root = root
        self.root.title("Downsview Sharer")

        tk.Label(root, text="This program will process spreadsheets (.csv or .xlsx) for the Keep@Downsview SHARE workflow.\nInputs must be valid spreadsheets with columns for at least uOttawa MMS ID, uOttawa Item PID, and UofT MMS ID.\n\nThis process has several steps:\n1. Current barcodes for all items associated with given PIDs will checked for Annex-type barcodes via API. These will be excluded from further processing. Items on loan or otherwise in processes will also be excluded.\n2. The given UofT MMS IDs will be checked using their Z39.50 portal for 'REMOTEOS' in a 949 $l. Records without a 949 $l REMOTEOS will be excluded from further processing.\n3. uOttawa's item records will have their barcodes stripped and locations updated to Downsview, as well as having item-level retention information added.\n4. New holdings records will be retrieved for uOttawa using the item PIDs (in case any holdings have been split across locations).\n5. Holdings records will be checked for updated location and have UofT's MMS IDs inserted in 014 fields\n6. uOttawa's bib records will have 940 fields with retention information added.\n").grid(row=0,column=0,columnspan=3,pady=10,padx=10)

        tk.Label(root, text="Input desired API Key. NOTE: Use a Sandbox key if wanting to test in the sandbox").grid(row=1,column=0,columnspan=3,pady=10,padx=10)
        self.apikey = tk.StringVar()
        self.apikeyentry = ttk.Entry(root,width=40)
        self.apikeyentry.grid(row=2,column=1,pady=10,padx=10)

        #Prompt user to load file
        tk.Label(root, text="STEP 1: Select an Excel/CSV file that includes the relevant fields (uOttawa MMS ID, Item PID, and UofT MMS ID)").grid(row=3,column=1,pady=10,padx=10)
        tk.Button(root, text="Load", command=self.getfile).grid(row=4,column=1,pady=10,padx=10)

        #Use column headers to select data for loading
        tk.Label(root, text="STEP 2: Select the column header that includes the uOttawa MMS ID, the uOttawa Item PID, and the UofT MMS ID").grid(row=5,column=1,pady=10,padx=10)

        self.uomms_box = tk.Listbox(root, exportselection=0, height=5)
        self.itemid_box = tk.Listbox(root, exportselection=0, height=5)
        self.uoftmms_box = tk.Listbox(root, exportselection=0, height=5)

        tk.Label(root, text="uOttawa MMS ID").grid(row=6,column=0,pady=10,padx=10)
        self.uomms_box.grid(row=7,column=0,pady=10,padx=10)
        tk.Label(root, text="Item PID").grid(row=6,column=1,pady=10,padx=10)
        self.itemid_box.grid(row=7,column=1,pady=10,padx=10)
        tk.Label(root, text="UofT MMS ID").grid(row=6,column=2,pady=10,padx=10)
        self.uoftmms_box.grid(row=7,column=2,pady=10,padx=10)

        #Run main function for complete workflow
        tk.Label(root, text="Step 3: Press below to run the complete process as described above. Note that the process may take some time to complete, so check the terminal for status updates.").grid(row=8,column=0,columnspan=3,pady=10,padx=10)
        tk.Button(root, text="RUN", command=self.runsharer).grid(row=9,column=1,pady=10,padx=10)

        #Create dataframe and input check for validation
        self.df = None
        self.inputpresent = None
        self.uoerrorlist = ["MMS Id"]
        self.annexitems = ["Item PID"]
        self.itemsonloan = ["Item PID"]
        self.itemnotfound = ["Item PID"]
        self.itemsinstatus = ["Item PID"]
        self.itemupdatefailed = ["Item PID"]
        self.holdingsupdatefailed = ["Holdings Id"]
        self.bibupdatefailed = ["MMS Id"]
        self.bibnotfound_uoft = ["MMS Id"]
        self.unsharedrecs = ["MMS Id"]
        self.today = datetime.datetime.today().strftime('%Y/%m/%d')
        self.currentholdings = ""
        self.updatedbibs = []

    def getfile(self):
            self.uomms_box.delete(0, tk.END)
            self.itemid_box.delete(0, tk.END)
            self.uoftmms_box.delete(0, tk.END)
            file = filedialog.askopenfilename()
            #Change working directory to folder of source file
            os.chdir(file[:file.rfind("/")])
            #Check filetype and populate dataframe
            if file.endswith("xlsx") or file.endswith("xlsm") or file.endswith("xls"):
                try:
                    self.df = pd.read_excel(file)
                    print("XLSX file recognized and read")
                    for column in list(self.df.columns.values):
                        self.uomms_box.insert(tk.END, column)
                        self.itemid_box.insert(tk.END, column)
                        self.uoftmms_box.insert(tk.END, column)
                    self.inputpresent = True
                    return self.inputpresent
                except Exception as e:
                    messagebox.showerror(title="Fileread error - XLSX", message=e)
                    self.inputpresent = False
                    return self.inputpresent
            elif file.endswith("csv"):
                try:
                    self.df = pd.read_csv(file)
                    print("CSV file recognized and read")
                    for column in list(self.df.columns.values):
                        self.uomms_box.insert(tk.END, column)
                        self.itemid_box.insert(tk.END, column)
                        self.uoftmms_box.insert(tk.END, column)
                    self.inputpresent = True
                    return self.inputpresent
                except Exception as e:
                    messagebox.showerror(title="Fileread error - CSV", message=e)
                    self.inputpresent = False
                    return self.inputpresent
            else:
                messagebox.showerror(title="Filetype Error", message="Error: Please ensure file is in .xlsx or .csv format and filename includes filetype extension.")
                self.inputpresent = False
                return self.inputpresent

    def runsharer(self):
        #Clear error lists to avoid duplication
        self.uoerrorlist = ["MMS Id"]
        self.annexitems = ["Item PID"]
        self.itemsonloan = ["Item PID"]
        self.itemnotfound = ["Item PID"]
        self.itemsinstatus = ["Item PID"]
        self.itemupdatefailed = ["Item PID"]
        self.holdingsupdatefailed = ["Holdings Id"]
        self.bibupdatefailed = ["MMS Id"]
        self.bibnotfound_uoft = ["MMS Id"]
        self.unsharedrecs = ["MMS Id"]
        #Check for input and get data for running processes
        if self.df.empty:
            messagebox.showerror(title="Input Not Found", detail="Error: Please ensure a valid input file has been selected.")
            return None
        #Get only relevant columns in dataframe
        uomms_column = self.uomms_box.get(self.uomms_box.curselection())
        itemid_column = self.itemid_box.get(self.itemid_box.curselection())
        uoftmms_column = self.uoftmms_box.get(self.uoftmms_box.curselection())
        #Check that all columns have input and none are duplicates
        if uomms_column == None or itemid_column == None or uoftmms_column == None:
            messagebox.showerror(title="Missing Inputs",message="Error: Please select a column for every input")
            return None
        if uomms_column == itemid_column or itemid_column == uoftmms_column or uomms_column == uoftmms_column:
            messagebox.showerror(title="Error: Duplicate Columns", message="Error: Please ensure correct columns are selected for each data element. Selecting the same column header multiple times is not allowed.")
            return None
        self.df = self.df[[uomms_column,itemid_column,uoftmms_column]]
        #Validate API Key input
        apivalidator = re.compile(r'\w+')
        self.apikey = self.apikeyentry.get().strip()
        if not re.match(apivalidator,self.apikey):
            messagebox.showerror(title="Error: Invalid API Key", message="Error: Please ensure the API key entered is correct and contains only alphanumeric characters in a continuous string.")
            return None
        headers = {
                        "Authorization" : f'apikey {self.apikey}',
                        "Accept" : "application/xml",
                        "Content-Type" : "application/xml"
                    }
        #For each row in spreadsheet, run complete process, updating outputs as needed.
        for index, row in self.df.iterrows():
            uomms = row[uomms_column]
            itemid = row[itemid_column]
            uoftmms = row[uoftmms_column]
            bcode_ok = False
            bcode_ok = self.barcodecheck(uomms,itemid,headers)
            if bcode_ok != True:
                print(f'Barcode check failed for item {itemid}. Continuing...')
                continue
            uoft_ok = False
            uoft_ok = self.uoftmmsget(uoftmms)
            if uoft_ok != True:
                print(f'UofT MMS check failed for record {uoftmms}. Continuing...')
                continue
            itemupdate_ok = False
            itemupdate_ok = self.itemupdate(uomms, itemid, headers)
            if itemupdate_ok != True:
                print(f'Item update for item {itemid} failed. Continuing...')
                continue
            holdingsupdate_ok = False
            holdingsupdate_ok = self.holdingsupdate(uomms, uoftmms, itemid, headers)
            if holdingsupdate_ok != True:
                print(f'Holdings update failed for holdings record associated with item {itemid}. Continuing...')
                continue
            bibupdate_ok = False
            bibupdate_ok = self.bibupdate(uomms, headers)
            if bibupdate_ok != True:
                print(f'Update for bibliographic record {uomms} failed. Continuing...')
                continue
        #Write outputs to directory of sourcefile using timestamp in filename to avoid overwriting
        #Set timestamp
        timestamp = datetime.datetime.today().strftime('%Y%m%d_%H%M')
        currdir = os.getcwd()
        for i in [(f'{timestamp} - uOttawa Holdings List Retrieval Failed.txt',self.uoerrorlist),(f'{timestamp} - Items with Annex Barcodes.txt',self.annexitems),(f'{timestamp} - Items on Loan.txt',self.itemsonloan),(f'{timestamp} - uOttawa Items Not Found.txt',self.itemnotfound),(f'{timestamp} - uOttawa Items in Status.txt',self.itemsinstatus),(f'{timestamp} - uOttawa Item Update Failed.txt',self.itemupdatefailed),(f'{timestamp} - uOttawa Holdings Update Failed.txt',self.holdingsupdatefailed),(f'{timestamp} - uOttawa Bib Update Failed.txt',self.bibupdatefailed),(f'{timestamp} - UofT Bib Not Found.txt',self.bibnotfound_uoft),(f'{timestamp} - UofT Bib Missing 949 REMOTEOS.txt',self.unsharedrecs)]:
            if len(i[1]) > 1:
                with open(i[0], "w") as fh:
                    fh.write("\n".join(i[1]))
                    print(f'{i[0]} saved to {os.getcwd()}')
            else:
                print(f'No content for {i[0]}. Continuing...')
        #Display completed message
        messagebox.showinfo(title="Process Complete!", message="Process completed successfully and relevant outputs saved!")

    def barcodecheck(self, uomms, itemid, headers):
        ###CAN ADJUST TO CHECK VALIDITY OF CURRENT BARCODES AS WELL? COULD BE HANDY FOR ERROR-CHECKING, BUT BCODES WILL BE STRIPPED ANYWAY, SO MAY NOT BE NECESSARY.
        #Set regex for catching annex-type barcodes
        annexbcode = re.compile(r'^33.*')
        print("Retrieving holdings...")
        #Spit out if bib on uOttawa errorlist
        if uomms in self.uoerrorlist:
            print("Bib already failed retrieval. Skipping...")
            return False
        #Get list of holdings for bib
        holdingslist = None
        for i in range(7):
            try:
                holdingslist = requests.get(f'{BASEURL}bibs/{uomms}/holdings',headers=headers)
                print("Holdings list retrieval successful!")
                break
            except Exception as e:
                print(e)
                print(f'Error retrieving holdings. Retrying in {2**i} seconds...')
                time.sleep(2**i)
                continue
        if holdingslist == None:
            print(f'Holdings retrieval failed too many times. Adding record {uomms} to error list and removing from processing.')
            self.uoerrorlist.append(str(uomms))
            return False
        #Parse holdingslist
        holdingslist = ET.fromstring(holdingslist.content)
        #Check all holdings for target item PID
        holdingsids = []
        for holdingsid in holdingslist.findall(".//holding_id"):
            holdingsids.append(str(holdingsid.text))
        itemfound = False
        for holdingsid in holdingsids:
            for i in range(7):
                if itemfound == False:
                    try:
                        item = requests.get(f'{BASEURL}bibs/{uomms}/holdings/{holdingsid}/items/{itemid}',headers=headers)
                        if item.status_code == 200:
                            itemfound = True
                            print("Item found!")
                    except Exception as e:
                        print(f'Error: {e}. Waiting {2**i} seconds before retrying...')
                        time.sleep(2**i)
                        continue
                else:
                    break
        #Add item to missing list if not found attached to any holdings of associated bib
        if itemfound == False:
            print(f'Item {itemid} not found. Adding to missing items list and continuing.')
            self.itemnotfound.append(str(itemid))
            return False
        #Parse item record for further error checking
        item = ET.fromstring(item.content)
        #Check for annex barcodes
        for bcode in item.findall(".//barcode"):
            if bcode.text == None:
                print("Barcode already removed!")
                break
            if re.match(annexbcode,bcode.text):
                print(f'Annex-type barcode {bcode} located. Adding to list of annex items and continuing.')
                self.annexitems.append(str(itemid))
                return False
        #Check for items on loan or otherwise in process
        for status in item.findall(".//process_type"):
            if status.text == "LOAN":
                print(f'Item {itemid} currently on-loan. Adding to list of loaned items and continuing.')
                self.itemsonloan.append(str(itemid))
                return False
            elif status.text != "LOAD" and status.text != None:
                print(f'Item {itemid} in status. Adding to list of in-status items and continuing.')
                self.itemsinstatus.append(str(itemid))
                return False
            else:
                print(f'Item {itemid} valid for processing. Continuing...')
        self.currentholdings = holdingsid
        return True

    def uoftmmsget(self, uoftmms):
        for i in range(7):
            try:
                conn = zoom.Connection("utoronto.alma.exlibrisgroup.com", 1921)
                conn.databaseName = "01UTORONTO_INST"
                conn.preferredRecordSyntax = "USMARC"
                #Run a raw query with @attr 1=12 (local number) and the MMS
                query = zoom.Query('PQF',f'@attr 1=12 {uoftmms}')
                res = conn.search(query)
                #Check for single record returned ONLY
                if len(res) > 1:
                    print(f'Error: Multiple records found for MMS {uoftmms}. Adding to list of unfound UofT records and continuing...')
                    self.bibnotfound_uoft.append(str(uoftmms))
                    conn.close()
                    return False
                elif len(res) < 1:
                    print(f'Error: No record found for MMS {uoftmms}. Adding to list of unfound UofT records and continuing...')
                    self.bibnotfound_uoft.append(str(uoftmms))
                    conn.close()
                    return False 
                #Take string of output and check for REMOTEOS using regex search
                remoteos = re.compile(r'949.*\$lREMOTEOS')
                if re.search(remoteos,str(res[0])):
                    print(f'K@D SHARE marking found in record. Continuing...')
                    conn.close()
                    return True
                else:
                    print(f'K@D SHARE marking not found in UofT record. Adding record to list of unshared UofT records and continuing...')
                    self.unsharedrecs.append(str(uoftmms))
                    conn.close()
                    return False
            except Exception as e:
                print(f'Error: {e}. Closing connection and waiting {2**i} seconds before trying again...')
                conn.close()
                time.sleep(2**i)
                continue

    def itemupdate(self, uomms, itemid, headers):
        #Get item using correct holdings from earlier in process
        for i in range(7):
            try:
                item = requests.get(f'{BASEURL}bibs/{uomms}/holdings/{self.currentholdings}/items/{itemid}',headers=headers)
                break
            except Exception as e:
                print(f'Error: {e}. Item retrieval failed. Waiting {2**i} seconds before retrying...')
                time.sleep(2**i)
                continue
        #Ensure item is correctly retrieved; mark as error if not
        if item.status_code == 200:
            item = ET.fromstring(item.content)
        else:
            print(f'Error getting item {itemid}. Adding to list of items not found and continuing.')
            self.itemnotfound.append(str(itemid))
            return False
        for i in item.findall(".//item_data"):
            for j in i.findall("./barcode"):
                j.text = ""
            for j in i.findall("./library"):
                j.attrib["desc"] = "Off-Campus Storage"
                j.text = "OFFCAMPUS"
            for j in i.findall("./location"):
                j.attrib["desc"] = "Storage - Downsview"
                j.text = "xdownsview"
            for j in i.findall("./committed_to_retain"):
                j.attrib["desc"] = "Yes"
                j.text = "true"
            for j in i.findall("./retention_reason"):
                j.attrib["desc"] = "Shared Print Collection - Keep@Downsview (uO)"
                j.text = "SharedPrintCollection-Keep@Downsview(uO)"
            for j in i.findall("./retention_note"):
                j.text = "2017/02/21 - 2037/02/21; Type: SHARE"
        for i in range(7):
            try:
                itemupdated = requests.put(f'{BASEURL}bibs/{uomms}/holdings/{self.currentholdings}/items/{itemid}',headers=headers,data=ET.tostring(item))
                if itemupdated.status_code == 200:
                    print(f'Item {itemid} correctly updated (barcode stripped, location changed, retention information added).')
                    return True
                else:
                    print(f'Item update of {itemid} failed with status code {itemupdated.status_code}. Waiting {2**i} seconds before retrying...')
                    time.sleep(2**i)
                    continue
            except Exception as e:
                print(f'Error: {e}. Waiting {2**i} seconds before retrying...')
                time.sleep(2**i)
                continue
        self.itemupdatefailed.append(str(itemid))
        return False        

    def holdingsupdate(self, uomms, uoftmms, itemid, headers):
        #Get list of holdings for bib
        holdingslist = None
        for i in range(7):
            try:
                holdingslist = requests.get(f'{BASEURL}bibs/{uomms}/holdings',headers=headers)
                print("Holdings list retrieval successful!")
                break
            except Exception as e:
                print(e)
                print(f'Error retrieving holdings. Retrying in {2**i} seconds...')
                time.sleep(2**i)
                continue
        #Parse holdingslist
        holdingslist = ET.fromstring(holdingslist.content)
        #Check all holdings for target item PID
        holdingsids = []
        for holdingsid in holdingslist.findall(".//holding_id"):
            holdingsids.append(str(holdingsid.text))
        correctholdings = ""
        for holdingsid in holdingsids:
            holdingsfound = False
            for i in range(7):
                if holdingsfound == False:
                    try:
                        item = requests.get(f'{BASEURL}bibs/{uomms}/holdings/{holdingsid}/items/{itemid}',headers=headers)
                        if item.status_code == 200:
                            correctholdings = holdingsid
                            print("Item found!")
                            holdingsfound = True
                    except Exception as e:
                        print(f'Error: {e}. Waiting {2**i} seconds before retrying...')
                        time.sleep(2**i)
                        continue
        #Get holdings record for editing
        for i in range(7):
            try:
                holdingsrec = requests.get(f'{BASEURL}bibs/{uomms}/holdings/{correctholdings}',headers=headers)
                break
            except Exception as e:
                print(f'Error: {e}. Waiting {2**i} seconds before retrying...')
                time.sleep(2**i)
                continue
        holdingsrec = ET.fromstring(holdingsrec.content)
        #Skip update if 014 with UofT MMS already in record
        for i in holdingsrec.findall(".//datafield[@tag='014']/subfield[@code='a']"):
            if i.text == str(uoftmms):
                print("Holdings record already updated. Continuing...")
                return True
        newholdings = self.updatexml(holdingsrec, uoftmms)
        for i in range(7):
            try:
                print(f'##################\nURL FOR HOLDINGS UPDATE: {f'{BASEURL}bibs/{uomms}/holdings/{correctholdings}'}\n#################')
                holdingsput = requests.put(f'{BASEURL}bibs/{uomms}/holdings/{correctholdings}',headers=headers,data=ET.tostring(newholdings))
            except Exception as e:
                print(f'Error: {e}. Waiting {2**i} seconds before retrying...')
                time.sleep(2**i)
                continue
            if holdingsput.status_code == 200:
                print("Holdings successfully updated. Continuing...")
                return True
            else:
                print(f'Holdings update unsuccessful with code {holdingsput.status_code}. Trying again after {2**i} seconds...')
                time.sleep(2**i)
                continue
        print(f'Update of holdings {correctholdings} failed too many times. Adding to list of failed updates and continuing...')
        self.holdingsupdatefailed.append(str(correctholdings))
        return False
                
    def updatexml(self, holdingsrecord, uoftmms):
        new014 = ET.Element("datafield")
        new014.set("tag","014")
        new014.set("ind1","1")
        new014.set("ind2"," ")
        suba = ET.Element("subfield")
        subb = ET.Element("subfield")
        suba.set("code","a")
        subb.set("code","b")
        suba.text = str(uoftmms)
        subb.text = "OTUTLD"
        new014.append(suba)
        new014.append(subb)
        for r in holdingsrecord.findall("record"):
            r.insert(1, new014)
        print(ET.tostring(holdingsrecord))
        return holdingsrecord

    def bibupdate(self, uomms, headers):
        if uomms in self.updatedbibs:
            print(f'Bib record {uomms} already updated! Continuing...')
            return True
        bibget = False
        for i in range(7):
            try:
                print(f'#################\nURL FOR BIB UPDATE IS {f'{BASEURL}bibs/{uomms}'}\n############')
                bibrecord = requests.get(f'{BASEURL}bibs/{uomms}',headers=headers)
                bibget = True
                break
            except Exception as e:
                print(f'Error: {e}. Waiting {2**i} seconds before retrying...')
                time.sleep(2**i)
                continue
        if bibget != True:
            print(f'Retrieval failed for bib {uomms}. Adding to error list and continuing...')
            self.bibupdatefailed.append(str({uomms}))
            return False
        bibrecord = ET.fromstring(bibrecord.content)
        #Check if bib is already updated with current date and workflow
        for i in bibrecord.findall(".//datafield[@tag='940']"):
            if i.find("./subfield[@code='a']").text == "Keep@Downsview" and i.find("./subfield[@code='b']").text == "SHARE" and i.find("./subfield[@code='c']").text == self.today:
                print("Bib record already updated. Continuing...")
                self.updatedbibs.append(uomms)
                return True
        new940 = ET.Element("datafield")
        new940.set("tag","940")
        new940.set("ind1"," ")
        new940.set("ind2"," ")
        suba = ET.Element("subfield")
        subb = ET.Element("subfield")
        subc = ET.Element("subfield")
        sub9 = ET.Element("subfield")
        suba.set("code","a")
        subb.set("code","b")
        subc.set("code","c")
        sub9.set("code","9")
        suba.text = "Keep@Downsview"
        subb.text = "SHARE"
        subc.text = f'{self.today}'
        sub9.text = "LOCAL"
        new940.extend([suba,subb,subc,sub9])
        for r in bibrecord.findall("record"):
            r.insert(1, new940)
        for i in range(7):
            try:
                bibput = requests.put(f'{BASEURL}bibs/{uomms}',headers=headers,data=ET.tostring(bibrecord))
            except Exception as e:
                print(f'Update of bib record {uomms} failed with exception {e}. Waiting {2**i} seconds and retrying...')
                time.sleep(2**i)
                continue
            if bibput.status_code == 200:
                print("Bib record successfully updated. Continuing...")
                self.updatedbibs.append(str(uomms))
                return True
            else:
                print(f'Bib update unsuccessful. Trying again after {2**i} seconds...')
                continue
        print(f'Update of bib record {uomms} failed too many times. Adding to list of failed updates and continuing...')
        self.bibupdatefailed.append(str(uomms))
        return False

def main():
    root = tk.Tk()
    app = DownsviewSharer(root)
    root.mainloop()

if __name__ == "__main__":
    main()