# Make sure that the locale 'en_SG.UTF-8' is installed!

from datetime import datetime, timedelta
import time
import pprint

import sqlitedb
import copyofresponses
import googlesheets
import kobo_import
import tabularise
from common import *
import sendmail

# Global Constants
KOBO_CREDENTIAL_FILE_NAME = 'kobo-credentials.json'
INHIBIT_FILE_NAME = 'inhibit'

# Global variables
debug = False
add_headers = False

# 'List of Repeats' XLS sheet
REPEATS_XLS_SHEET_NAME = 'List of repeats'
REPEATS_XLS_SHEET_START_COLUMN_CODE = 'A'
REPEATS_XLS_SHEET_START_COLUMN_OFFSET = 3
REPEATS_XLS_SHEET_STATUS_COLUMN_CODE = 'Q'

'''
Returns required configuration file name list.
'''
def getConfigFileList():
    return [KOBO_CREDENTIAL_FILE_NAME,
            GOOGLE_SHEET_IDS_FILE_NAME,
            GOOGLE_TOKENS_FILE_NAME,
            MAIL_TEMPLATE_STYLES_FILE_NAME,
            SMTP_CREDENTIALS_FILE_NAME]

'''
Perform environment checks to ensure all required settings are present.
'''
def _checkEnvironment():
    # Check for inhibit file for any errors in last run.
    if os.path.isfile(INHIBIT_FILE_NAME):
        printMsgAndQuit("Error: Tool is disabled due to error in last run")

    # Check all required configuration files present and connected to internet.
    checkEnvironment(getConfigFileList(), isInternetCheckRequired=True)

'''
Returns valid datetime object from the input date string.
'''
def getValidSubmittedDate(dateStr):
    if not dateStr:
        print("\tWarning: Input date string is empty while performing getValidSubmittedDate() !")
        return None

    # Possible date formats
    dateFormatList = ["%d %b %y", "%Y-%m-%d"]

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
        print("\tWarning: Input args are None while performing numDaysDifference() !")
        return None

    try:
        return (dateTime1 - dateTime2).days
    except Exception as e:
        print("\tError: Exception caught while performing numDaysDifference(), error: %s" % str(e))
        return None

'''
Fetch all 'List of repeats' sheet records from google sheet.
'''
def fetchAllListOfRepeatsRecords(googleSheetId):
    try:
        records = googlesheets.read_cells(googleSheetId,
                                          REPEATS_XLS_SHEET_NAME,
                                          REPEATS_XLS_SHEET_START_COLUMN_CODE, REPEATS_XLS_SHEET_START_COLUMN_OFFSET,
                                          REPEATS_XLS_SHEET_STATUS_COLUMN_CODE)

    except Exception as e:
        print("\tError: Exception caught while reading 'List of repeats' records, error: %s" % str(e))
        return None

    return records

'''
Find latest repeat record for the input uniqueId.
Returns record,index tupple.
'''
def findLatestRepeatRecordForUniqueId(uniqueId, repeatRecords):
    if not uniqueId:
        return None, None

    # Search in reverve order
    lastIndex = len(repeatRecords) - 1
    for index in range(lastIndex, 0, -1):
        record = repeatRecords[index]
        if record and (len(record) > 2) and (record[1] == uniqueId):
            return record, (index + REPEATS_XLS_SHEET_START_COLUMN_OFFSET)

    return None, None

'''
Checks if input records have matching contact info.
'''
def isMatchingContactInfo(record1, record2):
    if (not record1 or (len(record1) < 14)) or (not record2 or (len(record2) < 14)):
        print("\tWarning: Bad Input objects while performing isMatchingContactInfo() !")
        return False

    isEmptyContactInfo = True # all contact info is empty
    isAtleastOneContactInfoMatching = False # atleast one contact info matches

    startContactIndex = 6 # Email
    endContactIndex = 13 # Nationality

    for index in range(startContactIndex, endContactIndex):
        isEmptyContactInfo = isEmptyContactInfo and (record1[index] == '') and (record2[index] == '')
        isAtleastOneContactInfoMatching = isAtleastOneContactInfoMatching or (not isEmptyContactInfo and (record1[index] == record2[index]))
        if isAtleastOneContactInfoMatching:
            return True

    return isEmptyContactInfo

'''
Checks if submission date matches between new and old record.
'''
def isSubmissionDataMatches(newRecord, oldRecord):
    if (not newRecord or (len(newRecord) < 1)) or (not oldRecord or (len(oldRecord) < 16)):
        print("\tWarning: Bad Input objects while performing isSubmissionDataMatches() !")
        return False

    newRecDateSubmittedStr = newRecord[0] # Date Submitted
    oldRecNextSubmissionStr = oldRecord[15] # Date of Next Submission

    newRecDateSubmittedDate = getValidSubmittedDate(newRecDateSubmittedStr)
    if not newRecDateSubmittedDate:
        return False

    oldRecNextSubmissionDate = getValidSubmittedDate(oldRecNextSubmissionStr)
    if not oldRecNextSubmissionDate:
        return False

    # Date difference should be in the range (0 -> 30) days
    numDaysDiff = numDaysDifference(newRecDateSubmittedDate, oldRecNextSubmissionDate)
    if (numDaysDiff != None) and (numDaysDiff >= 0) and (numDaysDiff <= 30):
        return True

    return False

'''
Update status of the repeat record for the input record index.
'''
def updateRepeatRecordStatus(googleSheetId, recordIndex, status):
    if not googleSheetId or (recordIndex < 0) or not status:
        print("\tWarning: Bad Input args while performing updateRepeatRecordStatus() !")
        return False

    try:
        return googlesheets.update_cell(googleSheetId,
                                        REPEATS_XLS_SHEET_NAME,
                                        REPEATS_XLS_SHEET_STATUS_COLUMN_CODE, recordIndex,
                                        status)
    except Exception as e:
        print("\tError: Exception caught while updating status for older repeat record (recordIndex:%d) (status:%s), error: %s" % (recordIndex, status, str(e)))

    return False

'''
Sends email with list of UniqueId which are identified for Manual Review.
'''
def sendEmailForRecordsMarkedManualReview(manualReviewUniqueIdList):
    if not manualReviewUniqueIdList or (len(manualReviewUniqueIdList) == 0):
        return False

    result = True

    sender_email = 'covidsgsurvey@washinseasia.org'
    receiver_emails = ['heiko@rothkranz.net', 'asa.immanuela@gmail.com']
    subject = "WISE COVID-19 SG Survey Manual Review Required"

    debugMsg("\tSending Manual Review Required emails to: %s" % ", ".join(receiver_emails))

    html = "<html><body><b>Manual review required for UniqueId:</b><ol>"
    for uniqueId in manualReviewUniqueIdList:
        html += "<li>" + uniqueId + "</li>"
    html += "</ol></body></html>"
    txt = ''

    try:
        smtp_credentials = loadJson(SMTP_CREDENTIALS_FILE_NAME)

        server = 'smtp.emaillabs.net.pl'
        port = 465
        smtp_conn = sendmail.connect_smtp(server, port, sendmail.ENCRYPTION_TLS, smtp_credentials['user'], smtp_credentials['password'])
        for receiver_email in receiver_emails:
            recipients = sendmail.send_email(smtp_conn, sender_email, receiver_email, subject, txt, html)

        if recipients is not None:
            debugMsg("\tSuccessfully send email to: %s" % receiver_email)
        else:
            print("\tError: Failed to send email to: %s" % receiver_email)
            result = False

        sendmail.disconnect_smtp(smtp_conn)

    except Exception as e:
        print("\tError: Exception caught while sending email, error: %s" % str(e))
        result = False

    return result

'''
Update 'List of repeats' records status based on new upload records.
Also sends email with list of UniqueId which requires Manual Review.
'''
def updateRepeatRecordsStatus(googleSheetId, oldRepeatRecords, newUploadRecords):
    if (not oldRepeatRecords or (len(oldRepeatRecords) == 0)) or (not newUploadRecords or (len(newUploadRecords) == 0)):
        return

    debugMsg("\nProcessing 'List of repeats' XLS sheet older records ...")
    debugMsg("New upload records count: %s" %len(newUploadRecords))
    debugMsg("Old repeat records count: %s" %len(oldRepeatRecords))

    manualReviewUniqueIdList = []

    for newRecord in newUploadRecords:
        # New record should contains atleast 14 elements
        if not newRecord or (len(newRecord) < 14):
            print("\tWARNING: Bad new record found !")
            continue # ignore bad new record

        # Find latest matching old record
        uniqueId = newRecord[1]
        matchingLatestOldRecord,recordIndex = findLatestRepeatRecordForUniqueId(uniqueId, oldRepeatRecords)

        # Check 'Have you submitted this form before?' matches between new and old records
        newRecSubmittedBefore = newRecord[2]
        if (newRecSubmittedBefore == 'Yes') and (matchingLatestOldRecord == None):
            manualReviewUniqueIdList.append(uniqueId + " (no matching old record)")
            continue
        if (newRecSubmittedBefore == 'No') and (matchingLatestOldRecord != None):
            manualReviewUniqueIdList.append(uniqueId + " (found previous record, but claims not to have submitted before)")
            continue
        if (newRecSubmittedBefore == 'No') and (matchingLatestOldRecord == None):
            continue # ignore, nothing to be done

        # If no matching old record found then ignore
        if not newRecSubmittedBefore:
            continue # ignore, nothing to be done

        # If matching old record is completed then ignore it
        if (len(matchingLatestOldRecord) >= 17) and (matchingLatestOldRecord[16] == 'Completed'):
            continue # ignore, nothing to be done

        # Check contact info matches between new and old records
        if not isMatchingContactInfo(newRecord, matchingLatestOldRecord):
            manualReviewUniqueIdList.append(uniqueId + " (contact info mismatch)")
            continue

        # Check submittion date matches between new and old records
        if not isSubmissionDataMatches(newRecord, matchingLatestOldRecord):
            manualReviewUniqueIdList.append(uniqueId + " (submission date mismatch)")
            continue

        # Mark old record as Completed.
        updateRepeatRecordStatus(googleSheetId, recordIndex, 'Completed')

    if (len(manualReviewUniqueIdList) > 0):
        sendEmailForRecordsMarkedManualReview(manualReviewUniqueIdList)

################################################################################

'''
Main function
'''
def main(argv):
    # Check environment
    _checkEnvironment()

    # Connect with SQLite database
    conn = sqlitedb.connect_db('kobo.db')
    if not conn: quitApp()

    # Update 'lasttime' in the database
    rowCount = sqlitedb.exec_sql(conn, f"UPDATE lastrun SET lasttime = { int(time.time()) }")
    if not rowCount: print("Error: Failed to update 'lasttime' in SQLite database")

    # Get last handled submission time
    last_submit_time = sqlitedb.exec_sql(conn, "SELECT lastsubmit FROM lastrun")[0][0]
    debugMsg(f'Last submit time: { last_submit_time }')

    try:
        '''Get the token from https://kf.kobotoolbox.org/token/
        Format for kobo-credentials.json:
        {
            "token": <token>
        }'''
        KOBO_TOKEN = loadJson(KOBO_CREDENTIAL_FILE_NAME)['token']

        kobo_import.init_kobo(KOBO_TOKEN)

        asset_uids = [
            'argHw9ZzcAtcmEytJbWQo7', # online version
            'aAYAW5qZHEcroKwvEq8pRb', # interview version
            'aXbWBpzEm8xZyatgciGnEd', # TWC2 version
        ]

        debugMsg('Kobo: Checking for new data')
        # Get new submissions since last handled submission time
        new_data = kobo_import.get_asset_data(asset_uids, last_submit_time)
        new_submissions = kobo_import.count_submissions(new_data)

        if not (new_submissions > 0 or add_headers):
            # Nothing to do, exit
            quitApp(errorCode=0)

        debugMsg('Kobo: Getting assets')
        assets = kobo_import.get_assets(asset_uids)

        debugMsg('Kobo: Extracting labels for questions, choices and answers')
        # Extract list of choices
        choice_lists = kobo_import.get_choice_lists(assets)
        if debug:
            print('choice_lists:')
            pprint.pprint(choice_lists)
        # Extract and merge questions
        questions_list = kobo_import.get_questions_list(assets)
        questions = kobo_import.get_merged_questions(questions_list)
        if debug:
            print('Merged questions:')
            pprint.pprint(questions)
        # Transform questions into "flat" tabular form
        flat_questions = tabularise.flatten_questions(questions)

        ######## INITIALISE GOOGLE ########
        # Push data to Google Sheets

        '''The Google Sheet ID show in the URL when you open a spreadsheet.
        Format for google-sheet-ids.json:
        {
            "CLEANEDDATA": <sheet_id>,
            "CONTACTS": <sheet_id>,
            "S70": <sheet_id>
        }'''
        GOOGLE_SHEET_IDS = loadJson(GOOGLE_SHEET_IDS_FILE_NAME)

        if add_headers:
            # Sheet is still empty -> fill the header rows first
            # Insert two new rows on top to preserve whatever may be in the sheet already
            debugMsg('Adding headers')

            sorted_values = tabularise.get_data_headers(flat_questions)
            contact_values = tabularise.get_contact_headers(flat_questions)
            for sheet_range in ['Data (labeled)', 'Data (codes)']:
                googlesheets.append_rows(GOOGLE_SHEET_IDS['CLEANEDDATA'], sheet_range, sorted_values)
            for sheet_range in ['Data (labeled)', 'Data (codes)']:
                googlesheets.append_rows(GOOGLE_SHEET_IDS['S70'], sheet_range, contact_values)
            exit()

        if new_submissions > 0:
            debugMsg(f'Labeling { new_submissions } submissions')
            # Concatenate the results from both versions and sort by submission time
            # Add the labels for questions and choices
            labeled_results = kobo_import.get_labeled_result(new_data, choice_lists, questions, questions_list)

            # Unpack and upload results
            for sheet_range in ['Data (labeled)', 'Data (codes)']:
                debugMsg(f'Uploading {new_submissions} submissions to {sheet_range}')
                if sheet_range == 'Data (labeled)':
                    value_type = 'labels'
                else:
                    value_type = 'codes'
                upload_values = tabularise.get_data_values(labeled_results, flat_questions, value_type)
                contact_upload_values = tabularise.get_contact_values(labeled_results, flat_questions, value_type)

                googlesheets.append_rows(GOOGLE_SHEET_IDS['CLEANEDDATA'], sheet_range, upload_values)
                googlesheets.append_rows(GOOGLE_SHEET_IDS['S70'], sheet_range, contact_upload_values)

            # Update partner list
            upload_values = tabularise.get_partner_values(labeled_results)
            if upload_values:
                googlesheets.append_rows(GOOGLE_SHEET_IDS['CONTACTS'], 'Interested partners', upload_values)

            # Send copy of responses
            copyofresponses.handle_copies(labeled_results, questions, flat_questions)

            # Fill the list of repeats
            # Get list of existing UniqueIDs
            uidlist = googlesheets.read_column(GOOGLE_SHEET_IDS['CLEANEDDATA'], 'Data (labeled)', COLUMNS['unique_id'], 4)
            if debug:
                print("List of UIDs:")
                pprint.pprint(uidlist)
            upload_values = tabularise.get_repeat_values(labeled_results, uidlist)
            if upload_values:
                repeatXLSSheetId = GOOGLE_SHEET_IDS['REPEAT_RESPONDENT']
                oldRepeatRecords = fetchAllListOfRepeatsRecords(repeatXLSSheetId)

                googlesheets.append_rows(GOOGLE_SHEET_IDS['CONTACTS'], 'List of repeats', upload_values)

                # Update 'List of repeats' records status based on newly fetched records.
                updateRepeatRecordsStatus(repeatXLSSheetId, oldRepeatRecords, upload_values)

            rowCount = sqlitedb.exec_sql(conn, f"UPDATE lastrun SET lastsubmit = '{ labeled_results[-1]['meta']['_submission_time'] }'")
            sqlitedb.disconnect_db(conn)

    except Exception as e:
        print("Error: Exception caught, error: %s" % str(e))
        # something went wrong -> disable the tool by creating a file named 'inhibit'
        open('inhibit', 'a').close()
        raise

if __name__ == '__main__':
    main(sys.argv[1:])
