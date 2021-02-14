# Make sure that the locale 'en_SG.UTF-8' is installed!

import time
import json
import pprint

import sqlitedb
import copyofresponses
import googlesheets
import kobo_import
import tabularise
from common import *

# Global Constants
KOBO_CREDENTIAL_FILE_NAME = 'kobo-credentials.json'
INHIBIT_FILE_NAME = 'inhibit'

# Global variables
debug = False
add_headers = False


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
        with open(KOBO_CREDENTIAL_FILE_NAME, 'r') as kobo_credentials_file:
            KOBO_TOKEN = json.load(kobo_credentials_file)['token']

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
        with open(GOOGLE_SHEET_IDS_FILE_NAME, 'r') as google_sheet_ids_file:
            GOOGLE_SHEET_IDS = json.load(google_sheet_ids_file)

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
                googlesheets.append_rows(GOOGLE_SHEET_IDS['CONTACTS'], 'List of repeats', upload_values)

            rowCount = sqlitedb.exec_sql(conn, f"UPDATE lastrun SET lastsubmit = '{ labeled_results[-1]['meta']['_submission_time'] }'")
            sqlitedb.disconnect_db(conn)

    except Exception as e:
        print("Error: Exception caught, error: %s" % str(e))
        # something went wrong -> disable the tool by creating a file named 'inhibit'
        open('inhibit', 'a').close()
        raise

if __name__ == '__main__':
    main(sys.argv[1:])
