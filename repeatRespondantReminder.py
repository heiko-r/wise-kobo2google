# Make sure that the locale 'en_SG.UTF-8' is installed!

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
XLS_DOCUMENT_NAME = 'REPEAT_RESPONDANT'
XLS_SHEET_NAME = 'List of repeats'

# XLS sheet column query data
XLS_SHEET_START_COLUMN_CODE = 'A'
XLS_SHEET_START_COLUMN_OFFSET = 3
XLS_SHEET_END_COLUMN_CODE = 'S'

# XLS sheet Status column data values to check
STATUS_CHECK_LIST = ['Send reminder', 'To submit today']


'''
Returns required configuration file name list.
'''
def getConfigFileList():
    return [GOOGLE_TOKENS_FILE_NAME]

'''
Perform environment checks to ensure all required settings are present.
'''
def _checkEnvironment():
    # Check all required configuration files present and connected to internet.
    checkEnvironment(getConfigFileList(), isInternetCheckRequired=True)

'''
Sends reminder email for the input record.
'''
def sendReminderEmail(record):
    result = True

    try:
        sender_email = 'covidsgsurvey@washinseasia.org'
        receiver_email = record[EMAIL_COLUMN_INDEX]
        subject = "Reminder email - Survey on COVID-19 behaviours in Singapore"

        debugMsg("\tSending reminder email to: %s" % receiver_email)

        # @todo: update below code
        #html = get_html(labeled_result, questions, flat_questions)
        #txt = get_txt(labeled_result, questions, flat_questions)
        html = "Sample html"
        txt = "sample text"

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
def updateRecord(record):
    try:
        debugMsg("\tUpdating record ...")

        #@todo - implement

    except Exception as e:
        print("\tError: Exception caught while updating record, error: %s" % str(e))
        return False

    return True

################################################################################

'''
Main function
'''
def main(argv):
    # Check environment
    _checkEnvironment()

    try:
        with open(GOOGLE_SHEET_IDS_FILE_NAME, 'r') as google_sheet_ids_file:
            GOOGLE_SHEET_IDS = json.load(google_sheet_ids_file)

        # Fetch records
        records = googlesheets.read_cells(GOOGLE_SHEET_IDS[XLS_DOCUMENT_NAME],
                                          XLS_SHEET_NAME,
                                          XLS_SHEET_START_COLUMN_CODE, XLS_SHEET_START_COLUMN_OFFSET,
                                          XLS_SHEET_END_COLUMN_CODE)
        recordCount = len(records)
        print("Total Records: %d" % recordCount)
        debugMsg("")

        # Stats
        reminderToSendCount = 0
        reminderSendSuccessCount = 0

        # Check all records
        for index in range(recordCount):
            record = records[index]
            recordLength = len(record)

            # Check for bad records
            if len(record) < (STATUS_COLUMN_INDEX + 1):
                debugMsg('\tWarning: Bad record detected (record length:%d). Record: %s' % (len(record), str(record)))
                continue

            # Check for email id presence
            if not record[EMAIL_COLUMN_INDEX]:
                continue

            # Check for records which require reminder
            if (record[STATUS_COLUMN_INDEX] in STATUS_CHECK_LIST) and (recordLength <= SECOND_FOLLOWUP_COLUMN_INDEX):
                reminderToSendCount += 1

                debugMsg(record)

                # Send reminder email and update rocord
                if sendReminderEmail(record) and updateRecord(record):
                    reminderSendSuccessCount += 1

        # Print stats
        debugMsg("")
        print('Reminders send: %d/%d (%d failed)' % (reminderSendSuccessCount, reminderToSendCount, (reminderToSendCount - reminderSendSuccessCount)))

    except Exception as e:
        print("Error: Exception caught while processing, error: %s" % str(e))
        quitApp() # return error to caller


if __name__ == '__main__':
    main(sys.argv[1:])
