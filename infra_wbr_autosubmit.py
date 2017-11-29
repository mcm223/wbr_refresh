# ##########
# WBR Automation Script
# NSDBI - mcgrathm@
# ##########
import win32com.client
import os
import time, datetime
import smtplib, ssl
import sys
import logging
from PyPDF2 import PdfFileMerger
from datetime import datetime as dt
from datetime import date, timedelta
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication

print "Import done."
time.sleep(5)

# Define week number and year for WBR
yr = dt.now().strftime("%Y")
wk = dt.now().strftime("%U")
wk = int(wk)-1
to_date = date.today()

print "Date setup done."
time.sleep(5)

# Set up error logging
logging.basicConfig(level=logging.DEBUG,
                    format='%(asctime)s %(levelname)-8s %(message)s',
                    datefmt='%a, %d %b %Y %H:%M:%S',
                    filename=os.path.normpath('D:/Users/mcgrathm/Desktop/scripts/logs/infra_log.log'),
                    filemode='w')
logging.debug('Script began running on ' + str(date.today()))

print "Logging setup done."
time.sleep(5)

# File location (old code)
#fileName="C:\Users\mcgrathm\Desktop\DEA_pWBR_Textbooks.xlsx"

# Load Excel file names (not full paths) into array from source directory
fileNames = []
fileNames = os.listdir("//ant/dept/AWSHWENG/ProductManagement/AWS NW WBR/infra_wbr_autosend/source_files/")

print "Files loaded."
time.sleep(5)

# Email recipients
#email_address_list = ['documentset+170@mega.email.amazon.com','nsdbi-infrawbr-autosubmit@amazon.com','infra-wbr-fpa@amazon.com','mcgrathm@amazon.com','jreisen@amazon.com'] #Prod version
email_address_list = ['mcgrathm@amazon.com'] #Dev version
error_email_address_list = ['mcgrathm@amazon.com'] #Reduce circulation if script errors out

print "Email addresses set up."
time.sleep(5)

# Method to build and send email
def send_mail(date,email_address_list,text,pdfPath):

    me = "mcgrathm@amazon.com"
    you = ", ".join(email_address_list)

    # Create message container 
    msg = MIMEMultipart('mixed') #Changed from alternative to mixed to fix iPhone bug

    msg['Subject'] = "NSD WBR Monitoring | InfraWBR Weekly Submission - %s" % (to_date)
    msg['From'] = "nsdbi_wbr_monitor-do-not-reply@amazon.com"
    msg['To'] = ", ".join(email_address_list)
    # msg['CC'] = "mcgrathm@amazon.com"

    print ("Creating text attachment...")
    # Record the MIME types of both parts - text/plain and text/html.
    part1 = MIMEText(text, 'html') # Should probably do two versions (plain and html) for compatibility
    msg.attach(part1)
    print ("Attached part 1")

    print ("Starting loop...")
    # Loop through files in pdfPath and attach them to the email
    for i in pdfPath:
        # Attach output PDF to email
        name = (os.path.normpath("C:/python_file_staging/") + "\\" + str(i))
        print name
        fp = open(name,'rb')
        part2 = MIMEApplication(fp.read(), _subtype = 'pdf')
        part2.add_header('content-disposition', 'attachment', filename = ('utf-8', '', i))
        msg.attach(part2)
        fp.close()
        print ("PDF Attached!")

    # Send the message via local SMTP server
    s = smtplib.SMTP('smtp.amazon.com')
    s.sendmail(me, email_address_list, msg.as_string()) 
    print('Sending email')
    s.quit()

# Method to iterate through file array and refresh/save docs
def process_deck(files):
    # Create file directory for weekly deck
    print ('Starting refresh...')

    # List of network printer ports
    ports = ["Ne00:", "Ne01:", "Ne02:", "Ne03:", "Ne04:","Ne05:", "Ne06:", "Ne07:", "Ne08:","Ne09:", "Ne10:", "Ne11:", "Ne12:","Ne13:", "Ne14:", "Ne15:", "Ne16:"]

    # Don't need to make new directory for this version
    # os.mkdir(os.path.normpath("C:/Users/mcgrathm/Desktop/wbr_output_folder/") + "\\" + str(wk) + "_" + str(yr))

    # Open Excel
    xl = win32com.client.DispatchEx("Excel.Application")
    xl.DisplayAlerts = False
    xl.Visible = False
    print ("Opening files...")

    # Interate file array to refresh/save
    for i in files:
        path = "//ant/dept/AWSHWENG/ProductManagement/AWS NW WBR/infra_wbr_autosend/source_files" + "/" + i
        print ("Opening file: " + path)
        wb = xl.workbooks.Open(path)
        time.sleep(7)
        xl.DisplayAlerts = False
        time.sleep(7)
        #Fixes bug with missing printer dialogue when no user logged in
        counter = -1
        while counter < 16:
            try:
                counter += 1
                xl.ActivePrinter = "Microsoft XPS Document Writer on " + ports[counter]
                print "Printer changed on port " + ports[counter]
            except Exception:
                pass
        time.sleep(7)
        print ("Refreshing file: " + path)
        wb.RefreshAll()
        time.sleep(15) # Sleep for a while to allow background refresh to complete. Without this, it'll save before refresh is done.
        print ("Calculating file...")
        xl.Calculate()
        time.sleep(7)
        print ("Saving file: " + path)
        wb.Save()
        #wb.WorkSheets(1).Select() #Select first sheet in document
        try:
            #wb.ActiveSheet.ExportAsFixedFormat(0,"test.pdf") #Saves to local documents folder
            wb.SaveAs(os.path.normpath("C:/python_file_staging/" + i[:-5] + ".pdf"), FileFormat=57)
            print ("Save successful!")
        except:
            # error handling
            print ("Error: Unable to convert")
        finally:
            xl.Workbooks.Close()
    else:
        xl.Quit()

# Method to combine PDFs from a file location into a single output
def merge_pdfs():
    # Paths for this new week that specify where the newly created pdf files are and where to output the combined pdf doc
    inputPath = os.path.normpath("C:/Users/mcgrathm/Desktop/wbr_output_folder/") + "\\" + str(wk) + "_" + str(yr)
    outputPath = os.path.normpath("C:/Users/mcgrathm/Desktop/wbr_output_folder/final_pdfs") 

    # Create directory for the week if none exists
    try:
        os.mkdir(outputPath)
        print("Directory created")
    except Exception:
        sys.exc_clear() # Ignore and clear error if dir already exists
        print("Direct already exists")

    # Load PDFs into array
    pdfs = []
    pdfs = os.listdir(inputPath)
    print pdfs

    # Create instance of file merger
    merger = PdfFileMerger()

    # Loop through array and add each PDF
    for i in pdfs:
        print i
        i = inputPath + "\\" +  i
        print i
        merger.append(i)

    # Write to an output PDF document
    writePath = outputPath + "\\" + str(wk) + "_" + str(yr) + ".pdf"
    merger.write(writePath)

# Save deck as PDF and send success/fail email
try:
    process_deck(fileNames)
    # merge_pdfs() # Not needed for single-file refresh script
    html = ("<p>This is an automated email, please do not reply. If you would like to add or revise slides, check out " +
            "<a href='https://w.amazon.com/bin/view/Networking/NetworkDeployment/Network_Scaling_and_Deployment_Business_Intelligence_(NSGBI)/InfraWBR_Auto-Submit_Service?'>" +
            "InfraWBR Auto-Submit Service</a> for details, or submit a change request here: <a href='https://tiny.amazon.com/xrcetoob/NSDBISIMRequest'>SIM Request</a>." +
            "<br><br>For general questions, contact <a href='mailto:mcgrathm@amazon.com'>mcgrathm@</a></p>")
    pdfPath = os.listdir(os.path.normpath("C:/python_file_staging/")) #Records an array of the file names within this directory
    print pdfPath
    print ("Sleeping. Sending email in:")
    for i in xrange(5,0,-1):
        time.sleep(1)
        sys.stdout.write(str(i)+' ')
        sys.stdout.flush()
    send_mail(to_date, email_address_list,html,pdfPath)
    logging.debug('Script completed successfully!')
except Exception:
    logging.error('Error!', exc_info=True) # Log the traceback 
finally:
    print "Success: Ending Happily"