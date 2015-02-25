import csv, sys, os, datetime, email, getpass, imaplib, time, subprocess, win32com.client

workDir = 'C:\\QBImport\\'
outputDir = workDir + "ready_for_import\\"
sourceDir = workDir + "processed_source\\"
todaysDate = datetime.datetime.now().strftime('%Y-%m-%d')
userName = "justin.carlson@worldofcarlson.com"
passwd = "Mickeymouse11"

#importArgs = "/TEXT_FILE=C:\\csv\\ready_for_import\\invoiceExport_4_20141219_PROCESSED.txt /DELIMITER=Tab /TXN_TYPE=Invoice /MAP_FILE=C:\\csv\\maps\\invoiceImportFields.dat /LOG_ERROR"
#qbImport =  

#qbmn = subprocess.Popen("C:\\Program Files (x86)\\Intuit\\QuickBooks Enterprise Solutions 15.0\\QBW32Enterprise.exe E:\\New\\MN.QBW", stdin=None, stdout=None, stderr=None) 

if 'ready_for_import' not in os.listdir(workDir):
    os.mkdir(workDir + 'ready_for_import')

if datetime.datetime.now().strftime('%Y-%m-%d') not in os.listdir(sourceDir):
    os.mkdir(sourceDir + todaysDate)

if 'processed_source' not in os.listdir(workDir):
    os.mkdir(workDir + "processed_source")

if datetime.datetime.now().strftime('%Y-%m-%d') not in os.listdir(outputDir):
    os.mkdir(outputDir + todaysDate)


def sanitizeInvoice(fin, fout):
    with open(fin, 'r') as file_in:
        with open(fout,'w') as file_out:
            csv_reader = csv.DictReader(file_in, delimiter='\t', quotechar='"')

            fieldnames = ['fName', 'lName', 'company', 'email', 'addr1', 'city', 'state', 'zip', 'terms', 
                          'apptDate', 'product', 'description', 'quantity', 'price', 
                          'total', 'taxAmt', 'invoiceNum']

            csv_writer = csv.DictWriter(file_out, delimiter='\t', fieldnames=fieldnames)
            csv_writer.writeheader()

            for row in csv_reader:
                row['description'] = row['description'].replace('/n','').replace(',','')
                row['product'] = row['product'][:31].strip()
                row['company'] = row['product'][:41].strip()
                csv_writer.writerow(row)
            print("Payment Data Sanitized")

def sanitizePayment(fin, fout):
    with open(fin, 'r') as file_in:
        with open(fout,'w') as file_out:
            csv_reader = csv.DictReader(file_in, delimiter='\t', quotechar='"')

            fieldnames = ['fName', 'lName', 'method', 'amt', 'apptDate', 'referenceNum']

            csv_writer = csv.DictWriter(file_out, delimiter='\t', fieldnames=fieldnames)
            csv_writer.writeheader()

            for row in csv_reader:
                if row['referenceNum'] == "":
                    row['referenceNum'] = "cash"

                csv_writer.writerow(row)
            print("Payment Data Sanitized")

def downloadEmails():
    try:
        imapSession = imaplib.IMAP4_SSL('imap.gmail.com',993)
        typ, accountDetails = imapSession.login(userName, passwd)
        if typ != 'OK':
            print ('Not able to sign in!')
            raise

        imapSession.select('Inbox')
        typ, data = imapSession.search(None, 'ALL', 'UNSEEN', '(SUBJECT "Files for")')
        if typ != 'OK':
            print ('Error searching Inbox.')
            raise

        # Iterating over all emails
        for msgId in data[0].split():
            typ, messageParts = imapSession.fetch(msgId, '(RFC822)')

            if typ != 'OK':
                print ('Error fetching mail.')
                raise 

            emailBody = messageParts[0][1]
            mail = email.message_from_bytes(emailBody)

            for part in mail.walk():
                if part.get_content_maintype() == 'multipart':
                    continue
                if part.get('Content-Disposition') is None:
                    continue

                fileName = part.get_filename()

                if bool(fileName):
                    filePath = os.path.join(workDir, fileName)
                    if not os.path.isfile(filePath) :
                        print (fileName)
                        fp = open(filePath, 'wb')
                        fp.write(part.get_payload(decode=True))
                        fp.close()

        imapSession.close()
        imapSession.logout()

    except :
        print ('Not able to download all attachments.')
        time.sleep(3)

downloadEmails()


for file in os.listdir(workDir):

    print(file)
    if "invoice" in file:
        sanitizeInvoice(workDir + file, outputDir + "\\" + todaysDate + "\\" + file[:-10] + "PROCESSED.txt")
        os.rename(workDir + file, sourceDir + "\\" + todaysDate + "\\" + file)

    if "payment" in file:
        sanitizePayment(workDir + file, outputDir + "\\" + todaysDate + "\\" + file[:-10] + "PROCESSED.txt")
        os.rename(workDir + file, sourceDir + "\\" + todaysDate + "\\" + file)

    if "customer" in file:
        #Purge unused file
        os.remove(workDir + file)

print("Calling Quickbooks for import!")

#qbmn
# subprocess.check_call([qbImport, importArgs])