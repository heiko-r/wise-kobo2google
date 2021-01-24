# Make sure that the locale 'en_SG.UTF-8' is installed!

from datetime import datetime, timedelta
import googlesheets
import json
import pprint

from common import *
import sendmail

# Record based index (zero based indexes)
EMAIL_COLUMN_INDEX = 6
STATUS_COLUMN_INDEX = 16
FIRST_FOLLOWUP_COLUMN_INDEX = 17
SECOND_FOLLOWUP_COLUMN_INDEX = 18

# XLS sheet info
XLS_DOCUMENT_NAME = 'REPEAT_RESPONDENT'
XLS_SHEET_NAME = 'List of repeats'

# XLS sheet column query data
XLS_SHEET_START_COLUMN_CODE = 'A'
XLS_SHEET_START_COLUMN_OFFSET = 3
XLS_SHEET_END_COLUMN_CODE = 'S'

# XLS sheet followup column codes
XLS_FIRST_FOLLOWUP_COLUMN_CODE = 'R'
XLS_SECOND_FOLLOWUP_COLUMN_CODE = 'S'

# XLS sheet Status column data values to check
STATUS_CHECK_LIST = ['Send reminder', 'To submit today']

# Email content templates
MAIL_TEMPLATE_TXT_FILE_NAME = 'remindertemplate.txt'
MAIL_TEMPLATE_HTML_FILE_NAME = 'remindertemplate.html'

SECOND_REMINDER_EMAIL_DAY_DIFFERENCE = 5 # second reminder should be send only starting 5th day onwards

'''
Returns required configuration file name list.
'''
def getConfigFileList():
    return [GOOGLE_TOKENS_FILE_NAME]

'''
Returns todays date string in format 'mm/dd/yyyy'
'''
def getTodaysDateString():
    now = datetime.now()
    return now.strftime("%d/%m/%Y")

'''
Returns valid datetime object from the input text.
'''
def getValidDateTimeFromText(text):
    if not text:
        return None

    dateStr = text.split(' ')[-1]
    if not dateStr:
        return None

    # Possible date formats
    dateFormatList = ["%d/%m/%Y", "%d/%m/%y", "%d-%m-%Y", "%d-%m-%y"]

    # Check for valid date string pattern in the text
    for dateFormat in dateFormatList:
        try:
            return datetime.strptime(dateStr, dateFormat)
        except ValueError:
            pass # ignore exception

    return None

'''
Returns number of days difference in input datetime objects.
'''
def numDaysDifference(dateTime1, dateTime2):
    if not dateTime1 or not dateTime2:
        return None

    try:
        return (dateTime1 - dateTime2).days
    except Exception as e:
        return None

def isReminderSendAllowed(record, index):
    recordLength = len(record)

    # Ignore bad records
    if recordLength < (STATUS_COLUMN_INDEX + 1):
        debugMsg('\tWarning: Bad record detected (record length:%d). Record: %s' % (len(record), str(record)))
        return False

    # Check for email id presence
    if not record[EMAIL_COLUMN_INDEX]:
        return False # can't send reminder email

    # Check record status
    if (record[STATUS_COLUMN_INDEX] not in STATUS_CHECK_LIST):
        return False # no reminder send required

    # Check first reminder
    firstReminderColumnText = record[FIRST_FOLLOWUP_COLUMN_INDEX] if (recordLength > FIRST_FOLLOWUP_COLUMN_INDEX) else None
    if not firstReminderColumnText:
        return True # no reminders send yet

    # Check second reminder
    secondReminderColumnText = record[SECOND_FOLLOWUP_COLUMN_INDEX] if (recordLength > SECOND_FOLLOWUP_COLUMN_INDEX) else None
    if secondReminderColumnText:
        return False # all reminders are send already

    # Determine second reminder date (we should send second reminder only aftyer SECOND_REMINDER_EMAIL_DAY_DIFFERENCE days from the first reminder date)
    if firstReminderColumnText:
        firstRemiderSendDateTime = getValidDateTimeFromText(firstReminderColumnText)
        if not firstRemiderSendDateTime:
            print("\tError: 'First follow up' column data format is invalid (text:%s) (xlsRowIndex:%d)" % (firstReminderColumnText, index + XLS_SHEET_START_COLUMN_OFFSET))
            return False # skip this record

        numDaysDiff = numDaysDifference(datetime.now(), firstRemiderSendDateTime)
        if numDaysDiff and (numDaysDiff > 4):
            return True

    return False

'''
Perform environment checks to ensure all required settings are present.
'''
def _checkEnvironment():
    # Check all required configuration files present and connected to internet.
    checkEnvironment(getConfigFileList(), isInternetCheckRequired=True)

'''
Fetch records from google sheet.
'''
def fetchRecords(googleSheetId):
    # Fetch records
    records = googlesheets.read_cells(googleSheetId,
                                      XLS_SHEET_NAME,
                                      XLS_SHEET_START_COLUMN_CODE, XLS_SHEET_START_COLUMN_OFFSET,
                                      XLS_SHEET_END_COLUMN_CODE)

    return records

'''
Sends reminder email for the input record.
'''
def sendReminderEmail(record):
    result = True

    try:
        sender_email = 'covidsgsurvey@washinseasia.org'
        receiver_email = record[EMAIL_COLUMN_INDEX]
        subject = "👋 WISE COVID-19 SG Survey Reminder"

        debugMsg("\tSending reminder email to: %s" % receiver_email)

        # Read the TXT and HTML content from files
        # No templating needed
        with open(MAIL_TEMPLATE_HTML_FILE_NAME, 'r') as mail_template_html_file:
            html = mail_template_html_file.read()
        with open(MAIL_TEMPLATE_TXT_FILE_NAME, 'r') as mail_template_txt_file:
            txt = mail_template_txt_file.read()

        with open(SMTP_CREDENTIALS_FILE_NAME, 'r') as smtp_credentials_file:
            smtp_credentials = json.load(smtp_credentials_file)

        server = 'smtp.emaillabs.net.pl'
        port = 465
        smtp_conn = sendmail.connect_smtp(server, port, sendmail.ENCRYPTION_TLS, smtp_credentials['user'], smtp_credentials['password'])
        recipients = sendmail.send_email(smtp_conn, sender_email, receiver_email, subject, txt, html)

        if recipients is not None:
            debugMsg("\tSuccessfully send reminder email to: %s" % receiver_email)
        else:
            print("\tError: Failed to send reminder email to: %s" % receiver_email)
            result = False

        sendmail.disconnect_smtp(smtp_conn)

    except Exception as e:
        print("\tError: Exception caught while sending reminder email, error: %s" % str(e))
        return False

    return result

'''
Update record after reminder email send successfully.
'''
def updateRecord(googleSheetId, record, recordIndex):
    try:
        debugMsg("\tUpdating record ...")

        isFirstReminder = True if (len(record) == FIRST_FOLLOWUP_COLUMN_INDEX) else False

        # Update followup column value for the record
        xlsColumnCode = XLS_FIRST_FOLLOWUP_COLUMN_CODE if isFirstReminder else XLS_SECOND_FOLLOWUP_COLUMN_CODE
        xlsRowIndex = recordIndex + XLS_SHEET_START_COLUMN_OFFSET
        value = getTodaysDateString()

        result = googlesheets.update_cell(googleSheetId,
                                          XLS_SHEET_NAME,
                                          xlsColumnCode, xlsRowIndex,
                                          value)
        if result:
            return True

    except Exception as e:
        print("\tError: Exception caught while updating record, error: %s" % str(e))

    return False

################################################################################

'''
Main function
'''
def main(argv):
    # Check environment
    _checkEnvironment()

    try:
        # Load sheet ids
        with open(GOOGLE_SHEET_IDS_FILE_NAME, 'r') as google_sheet_ids_file:
            GOOGLE_SHEET_IDS = json.load(google_sheet_ids_file)

        googleSheetId = GOOGLE_SHEET_IDS[XLS_DOCUMENT_NAME]

        # Fetch records
        records = fetchRecords(googleSheetId);

        recordCount = len(records)
        print("Total Records: %d" % recordCount)
        debugMsg("")

        # Stats
        reminderToSendCount = 0
        reminderSendSuccessCount = 0

        # Check all records
        for index in range(recordCount):
            record = records[index]

            if not record:
                print("\tError: Empty record object for xlsRowIndex:%d" % (index + XLS_SHEET_START_COLUMN_OFFSET))
                continue

            # Check if reminder send is allowed
            if not isReminderSendAllowed(record, index):
                continue

            reminderToSendCount += 1
            debugMsg(record)

            # Send reminder email and update rocord
            if sendReminderEmail(record) and updateRecord(googleSheetId, record, index):
                reminderSendSuccessCount += 1

            break

        # Print stats
        debugMsg("")
        print('Reminders send: %d/%d (%d failed)' % (reminderSendSuccessCount, reminderToSendCount, (reminderToSendCount - reminderSendSuccessCount)))

    except Exception as e:
        print("Error: Exception caught while processing, error: %s" % str(e))
        quitApp() # return error to caller


if __name__ == '__main__':
    main(sys.argv[1:])
