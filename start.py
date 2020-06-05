# Make sure that the locale 'en_SG.UTF-8' is installed!

import os
import sys
from koboextractor import KoboExtractor
import sqlite3
import time
import pickle # for Google auth
import json
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.http import MediaIoBaseUpload
import string
from datetime import datetime, timedelta
import jinja2
import locale
import pdfkit
from io import BytesIO

debug = False

add_headers = False

if debug:
    print(f'Current path: { os.getcwd() }')
    import pprint

def columnToIndex(column_string):
    index = -1
    for i in range(0, len(column_string)): # 0,1,2
        index += (string.ascii_uppercase.index(column_string[-i-1]) + 1) * (26**i)
    return index

def getUniqueId(labeled_result):
    global GC_ID, QC_ID1, QC_ID2, QC_ID3, QC_ID4, QC_ID5, GC_ABOUT, QC_AGE, debug
    if debug: print('Trying to extract ID')
    try:
        id1 = str(labeled_result['questions'][GC_ID]['questions'][QC_ID1]['answer_code']).upper()
        id2 = "{0:0=2d}".format(int(labeled_result['questions'][GC_ID]['questions'][QC_ID2]['answer_code']))
        id3 = str(labeled_result['questions'][GC_ID]['questions'][QC_ID3]['answer_code'])
        id4 = str(labeled_result['questions'][GC_ID]['questions'][QC_ID4]['answer_code'])
        if QC_ID5 in labeled_result['questions'][GC_ID]['questions']:
            id5 = str(labeled_result['questions'][GC_ID]['questions'][QC_ID5]['answer_code'])
        else:
            id5 = str(labeled_result['questions'][GC_ABOUT]['questions'][QC_AGE]['answer_code'][-1])
        if debug: print(f'ID: { id1 + id2 + id3 + id4 + id5 }')
        uniqueId = id1 + id2 + id3 + id4 + id5
    except KeyError as err:
        if debug: print('KeyError in Unique ID', err)
        uniqueId = ''
    return uniqueId

def nextDate(submission_time_iso, weeks):
    return (datetime.fromisoformat(submission_time_iso) + timedelta(days=weeks*7, hours=8)).strftime('%A, %d %B')

conn = sqlite3.connect('kobo.db')
db = conn.cursor()

# Create table if it does not exist yet
db.execute("SELECT name FROM sqlite_master WHERE type='table' and name='lastrun'")
if not db.fetchall():
    print("Setting up database")
    db.execute("CREATE TABLE lastrun (lasttime int, lastsubmit text)")
    db.execute("INSERT INTO lastrun VALUES (0, '')")

db.execute(f"UPDATE lastrun SET lasttime = { time.time() }")

'''Get the token from https://kf.kobotoolbox.org/token/
Format for kobo-credentials.json:
{
    "token": "0b3bc87dbaa7ef82ad00411e791537581c409e48"
}'''
with open('kobo-credentials.json', 'r') as kobo_credentials_file:
    KOBO_TOKEN = json.load(kobo_credentials_file)['token']


kobo = KoboExtractor(KOBO_TOKEN, 'https://kf.kobotoolbox.org/api/v2', debug=debug)

#print('Kobo: Listing assets')
#assets = kobo.list_assets()
#asset_uid_online = assets['results'][0]['uid']
#print('UID of first asset: ' + assets['results'][0]['uid'])
asset_uid_online = 'argHw9ZzcAtcmEytJbWQo7'
asset_uid_interview = 'aAYAW5qZHEcroKwvEq8pRb'

if debug: print('Kobo: Checking for new data')
# Get last handled submission time
db.execute("SELECT lastsubmit FROM lastrun")
db_result = db.fetchall()
if not db_result:
    last_submit_time = None
else:
    last_submit_time = db_result[0][0]
if debug: print(f'Last submit time: { last_submit_time }')

# Get new submissions since last handled submission time
new_data_online = kobo.get_data(asset_uid_online, submitted_after=last_submit_time)
new_data_interview = kobo.get_data(asset_uid_interview, submitted_after=last_submit_time)

if new_data_online['count'] > 0 or new_data_interview['count'] > 0 or add_headers:
    if debug: print('Kobo: Getting assets')
    asset_online = kobo.get_asset(asset_uid_online)
    asset_interview = kobo.get_asset(asset_uid_interview)
    
    if debug: print('Kobo: Getting labels for questions and answers')
    # Create dict of of choice options in the form of choice_lists[list_name][answer_code] = answer_label
    choice_lists_online = kobo.get_choices(asset_online)
    choice_lists_interview = kobo.get_choices(asset_interview)
    # merge the choice lists into one, with the online version taking precedence
    # can ignore the sequence number here, because the order within a list should be unaffected by merging
    choice_lists = {**choice_lists_interview, **choice_lists_online}
    if debug:
        print('choice_lists:')
        pprint.pprint(choice_lists)
    
    questions_online = kobo.get_questions(asset=asset_online, unpack_multiples=True)
    questions_interview = kobo.get_questions(asset=asset_interview, unpack_multiples=True)
    
    # merge the questions into one.
    # sequence numbers will be colliding, hence the two dicts cannot be simply merged
    # questions_online takes precedence; only additional questions from questions_interview will be added
    questions = {}
    sequence = 0
    for question_group, question_group_dict in questions_online.items():
        questions[question_group] = {}
        questions[question_group]['label'] = question_group_dict['label']
        questions[question_group]['questions'] = {}
    for question_group, question_group_dict in questions_interview.items():
        if question_group not in questions:
            questions[question_group] = {}
            questions[question_group]['label'] = question_group_dict['label']
            questions[question_group]['questions'] = {}
    for question_group in questions:
        if question_group in questions_online:
            for question_code, question_code_dict in questions_online[question_group]['questions'].items():
                questions[question_group]['questions'][question_code] = {}
                questions[question_group]['questions'][question_code]['type'] = question_code_dict['type']
                questions[question_group]['questions'][question_code]['sequence'] = sequence
                if 'label' in question_code_dict:
                    questions[question_group]['questions'][question_code]['label'] = question_code_dict['label']
                if 'list_name' in question_code_dict:
                    questions[question_group]['questions'][question_code]['list_name'] = question_code_dict['list_name']
                sequence += 1
        if question_group in questions_interview:
            for question_code, question_code_dict in questions_interview[question_group]['questions'].items():
                if question_code not in questions[question_group]['questions']:
                    questions[question_group]['questions'][question_code] = {}
                    questions[question_group]['questions'][question_code]['type'] = question_code_dict['type']
                    questions[question_group]['questions'][question_code]['sequence'] = sequence
                    if 'label' in question_code_dict:
                        questions[question_group]['questions'][question_code]['label'] = question_code_dict['label']
                    if 'list_name' in question_code_dict:
                        questions[question_group]['questions'][question_code]['list_name'] = question_code_dict['list_name']
                    sequence += 1
    
    # Remove all questions without labels or of the following types
    # Note: Types in use to keep 'note', 'select_one', 'select_multiple', 'integer', 'text', '_or_other'
    delete_types = ['start', 'end', 'today', 'begin_group', 'end_group', 'calculate']
    for question_group, question_group_dict in questions.items():
        # The [] part is building a list of question_codes where the question type is in the above delete list
        for question_code in [question_code for question_code, question_dict in question_group_dict['questions'].items() if question_dict['type'] in delete_types]: del questions[question_group]['questions'][question_code]
        for question_code in [question_code for question_code, question_dict in question_group_dict['questions'].items() if 'label' not in question_dict]: del questions[question_group]['questions'][question_code]
    # delete empty question groups
    for question_group in [question_group for question_group, question_group_dict in questions.items() if not question_group_dict['questions']]: del questions[question_group]
    if debug:
        print('Filtered questions:')
        pprint.pprint(questions)
    
    ###### UPLOAD TO GOOGLE
    
    if debug: print('Initialising Google Sheets API')
    # Push data to Google Sheets
    
    '''The Google Sheet ID show in the URL when you open a spreadsheet.
    Format for google-sheet-ids.json:
    {
        "CLEANEDDATA": "1OQ-a9r17VW_y4AeISaLKutNLKBDRw2QePemfYsFrm4k",
        "CONTACTS": "1pMETsSB08C40_y_dCVGGRIxAGhMb5N0TQ53EoOzN0gg",
        "S70": "1zYWFfCYTLHicHxdd7AA2pIHcKbQUfRNMOl8Gk9GoINM"
    }'''
    with open('google-sheet-ids.json', 'r') as google_sheet_ids_file:
        GOOGLE_SHEET_IDS = json.load(google_sheet_ids_file)
    
    GOOGLE_SCOPES = [
        'https://www.googleapis.com/auth/spreadsheets',
        'https://www.googleapis.com/auth/drive'
    ]
    
    GOOGLE_DATETIME_FORMULA = '=(TEXT(LEFT(INDIRECT("A"&ROW()),10),"yyyy-mm-dd ")&TEXT(RIGHT(INDIRECT("A"&ROW()),8),"hh:mm:ss"))+8/24'
    GOOGLE_DATE_FORMULA = '=(TEXT(LEFT(INDIRECT("B"&ROW()),10),"yyyy-mm-dd "))'
    GOOGLE_VERSION_DICT = {
        'vhmDCKNckjxeWkSsysrmdk': 'v1',
        'vGzNDDcKCRUii3uYWX9eUD': 'v2',
        'vjRG97FSWLWGffyzekiZfy': 'v3',
        'vPKpE4cDvZFBCjW6qLj4XD': 'v4',
        'vPCHmyu7wSmfEzkjtSwtp9': 'v5',
        'veCAJ7XCQ6zwkjM2CUyh9b': 'v6',
        'vM23DK5cLuj69eRoYBVtZn': 'v7',
        'v2PeUCJPA4ybra67zSrzqG': 'v8',
        'v2Y4EQ5UnfzduEhHLkpd5q': 'v9',
        'vqLkQC2VUfE2Sju4daVLtj': 'v11',
        'vUhbRnabqzoDfAG7k3Lxqu': 'v12',
        'vByBFqUiFGNKN6xy3dVp2h': 'v13',
        'v7W8eK3RQMQvhgW9zPHozv': 'v14',
        'vgMEc8bjmok3NLwENyEjjX': 'v15',
        'vbp3gzGpKqrQBxxD9aRTCC': 'v17',
        'vf4jqJPWTasMrZRw5RZQ6H': 'v18',
        'vCV48cTDPfGkc8LqAFZ8CR': 'v19',
        'vnt9ehxAEEYr2adWCgVCd7': 'v20',
        'v6vd8mSArfBwJkHvNr39H6': 'v21',
        'vMri5MYvayCRfS6yrXGxhc': 'v22',
        'vDqpxGYyF28VCgWJRimDTW': 'v1-i',
        'vPzpWuwFFCobmzT6BxtBpi': 'v2-i',
        'vVbU7nhpFDTRrHdTWWhSzU': 'v23',
        'vQBQZaxKXmjfrd6EUv2AyW': 'v24',
    }
    GOOGLE_UNIQUEID_AFTER_GROUP = 'S60'
    GOOGLE_UNIQUEID_AFTER_QUESTION = 'ID5'
    
    GOOGLE_COLUMN_UNIQUEID = 'KV'
    GOOGLE_COLUMN_LATESTRESPONSE = 'KW'
    GOOGLE_LATEST_RESPONSE_FORMULA = f'=if(INDIRECT({GOOGLE_COLUMN_LATESTRESPONSE}$1&ROW())="","",max(arrayformula(if({GOOGLE_COLUMN_UNIQUEID}:{GOOGLE_COLUMN_UNIQUEID}=INDIRECT({GOOGLE_COLUMN_LATESTRESPONSE}$1&ROW()),row({GOOGLE_COLUMN_UNIQUEID}:{GOOGLE_COLUMN_UNIQUEID})))))'
    
    GC_CONTACT = 'S70'
    QC_VOLUNTEER = 'volunteer'
    QC_PARTNER = 'partner'
    QC_ORGNAME = 'partner_org'
    QC_EMAIL = 'email'
    QC_WHATSAPP = 'whatsapp'
    QC_TELEGRAM = 'telegram'
    QC_FACEBOOK = 'fb'
    QC_VIBER = 'viber'
    QC_SMS = 'sms'
    QC_CONTACT_OTHER = 'other'
    QC_COPY = 'copy'
    QC_AGAIN = 'again'
    QC_REMIND = 'remind'
    GC_ID = 'S60'
    QC_ID1 = 'ID1'
    QC_ID2 = 'ID2'
    QC_ID3 = 'ID3'
    QC_ID4 = 'ID4'
    QC_ID5 = 'ID5'
    GC_ABOUT = 'S40'
    QC_AGE = 'age'
    QC_NATIONALITY = 'nationality'
    GC_INTRO = 'S10'
    QC_BEFORE = 'S10Q05'
    QC_INTERVIEWER = 'interviewer'
    
    """Shows basic usage of the Sheets API.
    Prints values from a sample spreadsheet.
    """
    creds = None
    # The file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('googletoken.pickle'):
        with open('googletoken.pickle', 'rb') as token:
            if debug: print('Loading Google token from file')
            creds = pickle.load(token)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            if debug: print('Refreshing token')
            creds.refresh(Request())
        else:
            if debug: print('Loading Google secrets from file')
            # Download this credentials file from Google
            flow = InstalledAppFlow.from_client_secrets_file(
                'google-api-credentials.json', GOOGLE_SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('googletoken.pickle', 'wb') as token:
            if debug: print('Saving new token')
            pickle.dump(creds, token)
    
    service = build('sheets', 'v4', credentials=creds)
    
    # Call the Sheets API
    sheet = service.spreadsheets()
    #version_code = sheet.values().get(spreadsheetId=GOOGLE_SHEET_ID_CLEANEDDATA, range='D2').execute().get('values', [])
    
    # Get list of existing UniqueIDs for later
    uidresult = sheet.values().get(spreadsheetId=GOOGLE_SHEET_IDS['CLEANEDDATA'], range=f'Data (labeled)!{GOOGLE_COLUMN_UNIQUEID}4:{GOOGLE_COLUMN_UNIQUEID}').execute()
    uidlist = uidresult.get('values', [])
    if debug:
        print("List of UIDs:")
        pprint.pprint(uidlist)
    
    # Put all questions from all groups into one list and sort by occurrence in the survey
    all_questions = []
    for question_group_code, question_group_dict in questions.items():
        for question_code, question_dict in question_group_dict['questions'].items():
            if 'label' in question_dict:
                label = question_dict['label']
            else:
                label = question_code
            all_questions.append({
                'group_code': question_group_code,
                'question_code': question_code,
                'question_label': label,
                'sequence': question_dict['sequence']
            })
    sorted_questions = sorted(all_questions, key = lambda question: question['sequence'])
    
if add_headers:
    # Sheet is still empty -> fill the header rows first
    # Insert two new rows on top to preserve whatever may be in the sheet already
    if debug: print('Adding headers')
    
    sorted_codes = [None, None, None, None, None, None]
    sorted_labels = ['_submission_time', 'Submission date and time (computed)', 'Date (computed)', '_version_', 'Version (computed)', '_id']
    contact_codes = [None, None, None, None, None]
    contact_labels = ['_submission_time', 'Submission date and time (computed)', '_id', 'Unique ID (computed)', 'Interviewer']
    for question in sorted_questions:
        if question['group_code'] != GC_CONTACT:
            sorted_codes.append(f'{ question["group_code"] }/{ question["question_code"] }')
            sorted_labels.append(question['question_label'])
        else:
            contact_codes.append(f'{ question["group_code"] }/{ question["question_code"] }')
            contact_labels.append(question['question_label'])
        # Add the unique ID and row # of last response
        if question['group_code'] == GOOGLE_UNIQUEID_AFTER_GROUP and question['question_code'] == GOOGLE_UNIQUEID_AFTER_QUESTION:
            sorted_codes.append('')
            sorted_labels.append('Unique ID (computed)')
            sorted_codes.append('')
            sorted_labels.append('Row # of latest response of respondent')
    sorted_values = [
        sorted_codes,
        sorted_labels
    ]
    contact_values = [
        contact_codes,
        contact_labels
    ]
    request_body = {
        "range": 'Data (labeled)',
        "majorDimension": 'ROWS',
        'values': sorted_values
    }
    request = sheet.values().append(spreadsheetId=GOOGLE_SHEET_IDS['CLEANEDDATA'], range='Data (labeled)', valueInputOption='USER_ENTERED', insertDataOption='INSERT_ROWS', body=request_body)
    response = request.execute()
    request_body = {
        "range": 'Data (codes)',
        "majorDimension": 'ROWS',
        'values': sorted_values
    }
    request = sheet.values().append(spreadsheetId=GOOGLE_SHEET_IDS['CLEANEDDATA'], range='Data (codes)', valueInputOption='USER_ENTERED', insertDataOption='INSERT_ROWS', body=request_body)
    response = request.execute()
    request_body = {
        "range": 'Data (labeled)',
        "majorDimension": 'ROWS',
        'values': contact_values
    }
    request = sheet.values().append(spreadsheetId=GOOGLE_SHEET_IDS['S70'], range='Data (labeled)', valueInputOption='USER_ENTERED', insertDataOption='INSERT_ROWS', body=request_body)
    response = request.execute()
    request_body = {
        "range": 'Data (codes)',
        "majorDimension": 'ROWS',
        'values': contact_values
    }
    request = sheet.values().append(spreadsheetId=GOOGLE_SHEET_IDS['S70'], range='Data (codes)', valueInputOption='USER_ENTERED', insertDataOption='INSERT_ROWS', body=request_body)
    response = request.execute()
    exit()
    
if new_data_online['count'] > 0 or new_data_interview['count'] > 0:
    if debug: print(f'Labeling { new_data_online["count"] } online reults and { new_data_interview["count"] } interview results')
    # Concatenate the results from both versions and sort by submission time
    new_results = kobo.sort_results_by_time(new_data_online['results'] + new_data_interview['results'])
    # Add the labels for questions and choices
    
    labeled_results = []
    for result in new_results: # new_results is a list of list of dicts
        # Unpack answers to select_multiple questions
        labeled_results.append(kobo.label_result(unlabeled_result=result, choice_lists=choice_lists, questions=questions, unpack_multiples=True))
    
    # Add results data LABELS
    if debug: print(f'Uploading { len(labeled_results) } to Data (labeled)')
    upload_values = []
    contact_upload_values = []
    for labeled_result in labeled_results:
        if labeled_result['meta']['_version_'] in GOOGLE_VERSION_DICT:
            version = GOOGLE_VERSION_DICT[labeled_result['meta']['_version_']]
        else:
            version = 'NEW'
        result_values = [labeled_result['meta']['_submission_time'], GOOGLE_DATETIME_FORMULA, GOOGLE_DATE_FORMULA, labeled_result['meta']['_version_'], version, labeled_result['meta']['_id']]
        if GC_INTRO in labeled_result['questions'] and QC_INTERVIEWER in labeled_result['questions'][GC_INTRO]['questions']:
            interviewer = labeled_result['questions'][GC_INTRO]['questions'][QC_INTERVIEWER]['answer_code']
        else:
            interviewer = ''
        contact_values = [labeled_result['meta']['_submission_time'], GOOGLE_DATETIME_FORMULA, labeled_result['meta']['_id'], getUniqueId(labeled_result), interviewer]
        for question in sorted_questions:
            if question['group_code'] in labeled_result['questions'] and question['question_code'] in labeled_result['questions'][question['group_code']]['questions']:
                new_value = labeled_result['questions'][question['group_code']]['questions'][question['question_code']]['answer_label']
                if question['group_code'] != GC_CONTACT:
                    result_values.append(new_value)
                else:
                    contact_values.append(new_value)
            else:
                if question['group_code'] != GC_CONTACT:
                    result_values.append('')
                else:
                    contact_values.append('')
            # Add the unique ID and row # of last response
            if question['group_code'] == GOOGLE_UNIQUEID_AFTER_GROUP and question['question_code'] == GOOGLE_UNIQUEID_AFTER_QUESTION:
                result_values.append(getUniqueId(labeled_result))
                result_values.append(GOOGLE_LATEST_RESPONSE_FORMULA)
        upload_values.append(result_values)
        contact_upload_values.append(contact_values)
    
    #print('!')
    #print('.')
    #print('!')
    
    request_body = {
        "range": 'Data (labeled)',
        "majorDimension": 'ROWS',
        'values': upload_values
    }
    request = sheet.values().append(spreadsheetId=GOOGLE_SHEET_IDS['CLEANEDDATA'], range='Data (labeled)', valueInputOption='USER_ENTERED', insertDataOption='INSERT_ROWS', body=request_body)
    response = request.execute()
    
    request_body = {
        "range": 'Data (labeled)',
        "majorDimension": 'ROWS',
        'values': contact_upload_values
    }
    request = sheet.values().append(spreadsheetId=GOOGLE_SHEET_IDS['S70'], range='Data (labeled)', valueInputOption='USER_ENTERED', insertDataOption='INSERT_ROWS', body=request_body)
    response = request.execute()
    
    # Add results data CODES
    if debug: print(f'Uploading { len(labeled_results) } to Data (codes)')
    upload_values = []
    contact_upload_values = []
    for labeled_result in labeled_results:
        if labeled_result['meta']['_version_'] in GOOGLE_VERSION_DICT:
            version = GOOGLE_VERSION_DICT[labeled_result['meta']['_version_']]
        else:
            version = 'NEW'
        result_values = [labeled_result['meta']['_submission_time'], GOOGLE_DATETIME_FORMULA, GOOGLE_DATE_FORMULA, labeled_result['meta']['_version_'], version, labeled_result['meta']['_id']]
        if GC_INTRO in labeled_result['questions'] and QC_INTERVIEWER in labeled_result['questions'][GC_INTRO]['questions']:
            interviewer = labeled_result['questions'][GC_INTRO]['questions'][QC_INTERVIEWER]['answer_code']
        else:
            interviewer = ''
        contact_values = [labeled_result['meta']['_submission_time'], GOOGLE_DATETIME_FORMULA, labeled_result['meta']['_id'], getUniqueId(labeled_result), interviewer]
        for question in sorted_questions:
            if question['group_code'] in labeled_result['questions'] and question['question_code'] in labeled_result['questions'][question['group_code']]['questions']:
                new_value = labeled_result['questions'][question['group_code']]['questions'][question['question_code']]['answer_code']
                if question['group_code'] != GC_CONTACT:
                    result_values.append(new_value)
                else:
                    contact_values.append(new_value)
            else:
                if question['group_code'] != GC_CONTACT:
                    result_values.append('')
                else:
                    contact_values.append('')
            # Add unique ID and row # of last response
            if question['group_code'] == GOOGLE_UNIQUEID_AFTER_GROUP and question['question_code'] == GOOGLE_UNIQUEID_AFTER_QUESTION:
                result_values.append(getUniqueId(labeled_result))
                result_values.append(GOOGLE_LATEST_RESPONSE_FORMULA)
        upload_values.append(result_values)
        contact_upload_values.append(contact_values)
    
    request_body = {
        "range": 'Data (codes)',
        "majorDimension": 'ROWS',
        'values': upload_values
    }
    request = sheet.values().append(spreadsheetId=GOOGLE_SHEET_IDS['CLEANEDDATA'], range='Data (codes)', valueInputOption='USER_ENTERED', insertDataOption='INSERT_ROWS', body=request_body)
    response = request.execute()
    
    request_body = {
        "range": 'Data (codes)',
        "majorDimension": 'ROWS',
        'values': contact_upload_values
    }
    request = sheet.values().append(spreadsheetId=GOOGLE_SHEET_IDS['S70'], range='Data (codes)', valueInputOption='USER_ENTERED', insertDataOption='INSERT_ROWS', body=request_body)
    response = request.execute()
    
    # Update partner list
    upload_values = []
    for labeled_result in labeled_results:
        if GC_CONTACT in labeled_result['questions'] and QC_PARTNER in labeled_result['questions'][GC_CONTACT]['questions'] and labeled_result['questions'][GC_CONTACT]['questions'][QC_PARTNER]['answer_code'] == '01':
            if debug: print('Found new partner!')
            # Person wants to partner -> add to the partner list
            upload_row = [None] * (columnToIndex('U') + 1) # Data list for columns A to U
            
            timestring = labeled_result['meta']['_submission_time']
            timeobj = datetime.fromisoformat(timestring)
            upload_row[columnToIndex('A')] = (timeobj + timedelta(hours=8)).date().isoformat()
            
            upload_row[columnToIndex('F')] = getUniqueId(labeled_result)
            
            upload_row[columnToIndex('M')] = labeled_result['questions'][GC_CONTACT]['questions'][QC_PARTNER]['answer_label']
            
            try:
                upload_row[columnToIndex('N')] = labeled_result['questions'][GC_CONTACT]['questions'][QC_ORGNAME]['answer_code']
            except KeyError as err:
                if debug: print('KeyError in partner_org/N', err)
            
            try:
                upload_row[columnToIndex('O')] = labeled_result['questions'][GC_CONTACT]['questions'][QC_EMAIL]['answer_code']
            except KeyError as err:
                if debug: print('KeyError in email/O', err)
            
            try:
                upload_row[columnToIndex('P')] = labeled_result['questions'][GC_CONTACT]['questions'][QC_WHATSAPP]['answer_code']
            except KeyError as err:
                if debug: print('KeyError in whatsapp/P', err)
            
            try:
                upload_row[columnToIndex('Q')] = labeled_result['questions'][GC_CONTACT]['questions'][QC_TELEGRAM]['answer_code']
            except KeyError as err:
                if debug: print('KeyError in telegram/Q', err)
            
            try:
                upload_row[columnToIndex('R')] = labeled_result['questions'][GC_CONTACT]['questions'][QC_FACEBOOK]['answer_code']
            except KeyError as err:
                if debug: print('KeyError in facebook/R', err)
            
            try:
                upload_row[columnToIndex('S')] = labeled_result['questions'][GC_CONTACT]['questions'][QC_VIBER]['answer_code']
            except KeyError as err:
                if debug: print('KeyError in viber/S', err)
            
            try:
                upload_row[columnToIndex('T')] = labeled_result['questions'][GC_CONTACT]['questions'][QC_SMS]['answer_code']
            except KeyError as err:
                if debug: print('KeyError in sms/T', err)
            
            try:
                upload_row[columnToIndex('U')] = labeled_result['questions'][GC_CONTACT]['questions'][QC_CONTACT_OTHER]['answer_code']
            except KeyError as err:
                if debug: print('KeyError in other/U', err)
            
            upload_values.append(upload_row)
    
    if upload_values:
        request_body = {
            "range": 'Interested partners',
            "majorDimension": 'ROWS',
            'values': upload_values
        }
        request = sheet.values().append(spreadsheetId=GOOGLE_SHEET_IDS['CONTACTS'], range='Interested partners', valueInputOption='USER_ENTERED', insertDataOption='INSERT_ROWS', body=request_body)
        response = request.execute()
    
    #print('-')
    #print('/')
    #print('|')
    
    upload_values = []
    for labeled_result in labeled_results:
        # Send copy of responses
        if GC_CONTACT in labeled_result['questions'] and QC_COPY in labeled_result['questions'][GC_CONTACT]['questions'] and labeled_result['questions'][GC_CONTACT]['questions'][QC_COPY]['answer_code'] == '01':
            if debug: print('Copy requested!')
            
            def templatedebug(text):
                """Filter for debugging templates.
                
                Callable by Jinja2 templates.
                Usage, e.g.: {{ questions.S70.questions.age.answer_code | templatedebug }}
                """
                print(text)
                return ''
            
            template_loader = jinja2.FileSystemLoader('./')
            template_env = jinja2.Environment(
                loader = template_loader,
                autoescape = jinja2.select_autoescape(['html', 'xml'])
            )
            template_env.filters['debug'] = debug
            html_template = template_env.get_template('mailtemplate.html')
            txt_template = template_env.get_template('mailtemplate.txt')
            pdf_template = template_env.get_template('pdftemplate.html')
            
            with open('mailtemplate-styles.json', 'r') as jsonfile:
                styles = json.load(jsonfile)
            
            locale.setlocale(locale.LC_ALL, 'en_SG.UTF-8')
            
            next_date = nextDate(labeled_result['meta']['_submission_time'], int(labeled_result['questions']['S70']['questions']['again']['answer_code']))
            
            template_data = {
                'questions': questions,
                'sorted_questions': sorted_questions,
                'labeled_result': labeled_result,
                'next_date': next_date,
                'styles': styles
            }
            
            if QC_EMAIL in labeled_result['questions'][GC_CONTACT]['questions']:
                # Send email automatically
                html = html_template.render(template_data)
                txt = txt_template.render(template_data)
                # Assemble and send the email
                import smtplib, ssl
                from email.mime.text import MIMEText
                from email.mime.multipart import MIMEMultipart
                from email.utils import formatdate
                from email.utils import make_msgid
                
                sender_email = 'covidsgsurvey@washinseasia.org'
                receiver_email = labeled_result['questions'][GC_CONTACT]['questions'][QC_EMAIL]['answer_label']
                #receiver_email = 'heiko@rothkranz.net'
                '''Format of gmail-credentials.json:
                {
                    "user": "user@example.com",
                    "password": "safEpa55w0rd"
                }'''
                with open('gmail-credentials.json', 'r') as gmail_credentials_file:
                    gmail_credentials = json.load(gmail_credentials_file)
                #user = '40784a442d78cc'
                #password = '2c2ea9a8135d5c'
                #server = 'smtp.mailtrap.io'
                server = 'smtp.gmail.com'
                port = 465
                
                message = MIMEMultipart("alternative")
                message["Subject"] = "Your responses - Survey on COVID-19 behaviours in Singapore"
                message["From"] = sender_email
                message["To"] = receiver_email
                message["Date"] = formatdate(localtime = True)
                message["Message-ID"] = make_msgid(domain='washinseasia.org')
                
                part1 = MIMEText(txt, "plain")
                part2 = MIMEText(html, "html")
                message.attach(part1)
                message.attach(part2)
                
                if debug: print(f'Sending mail to { receiver_email }')
                
                context = ssl.SSLContext(ssl.PROTOCOL_TLSv1_2)
                with smtplib.SMTP_SSL(server, port, context=context) as server:
                #with smtplib.SMTP(server, port) as server:
                    #server.starttls(context=context)
                    server.login(gmail_credentials['user'], gmail_credentials['password'])
                    server.sendmail(
                        sender_email, receiver_email, message.as_string()
                    )
                if debug: print('Mail sent!')
            else:
                # no email address -> create PDF and add respondent to the list of copies to be sent
                if debug: print('Copy requested, but no email given!')
                pdfhtml = pdf_template.render(template_data)
                if debug:
                    pdfoptions = {}
                else:
                    pdfoptions = {'quiet': ''}
                pdfdata = pdfkit.from_string(pdfhtml, False, options=pdfoptions)
                fh = BytesIO(pdfdata)
                filename = f'{ getUniqueId(labeled_result) }.pdf'
                drive_service = build('drive', 'v3', credentials=creds)
                file_metadata = {
                    'name': filename,
                    'parents': ['1zzWDH0Iltzp6cU9jqacF9MmLR3NT2SSx']
                }
                media = MediaIoBaseUpload(fh, mimetype='application/pdf')
                file = drive_service.files().create(body=file_metadata, media_body=media, fields='id', supportsAllDrives=True).execute()
                if debug: print(f'File ID: { file.get("id") }')
        
        #print('\\')
        #print('-')
        #print('/')
        
        if [getUniqueId(labeled_result)] in uidlist or (GC_INTRO in labeled_result['questions'] and QC_BEFORE in labeled_result['questions'][GC_INTRO]['questions'] and labeled_result['questions'][GC_INTRO]['questions'][QC_BEFORE]['answer_code'] == '01') or (GC_CONTACT in labeled_result['questions'] and QC_AGAIN in labeled_result['questions'][GC_CONTACT]['questions'] and labeled_result['questions'][GC_CONTACT]['questions'][QC_AGAIN]['answer_code'] != '00'):
            # Respondent has answered before, claims to have answered before or wants to do the survey again
            if debug: print('Found (potential) repeat respondent!')
            # Add to repeat respondent list
            upload_row = [None] * (columnToIndex('U') + 1) # Data list for columns A to U
            
            timestring = labeled_result['meta']['_submission_time']
            timeobj = datetime.fromisoformat(timestring)
            day_of_submission = (timeobj + timedelta(hours=8)).date()
            upload_row[columnToIndex('A')] = day_of_submission.isoformat()
            
            upload_row[columnToIndex('B')] = getUniqueId(labeled_result)
            
            try:
                upload_row[columnToIndex('E')] = labeled_result['questions'][GC_CONTACT]['questions'][QC_AGAIN]['answer_label']
            except KeyError as err:
                if debug: print('KeyError in again/E and T', err)
            
            try:
                upload_row[columnToIndex('F')] = labeled_result['questions'][GC_CONTACT]['questions'][QC_REMIND]['answer_label']
            except KeyError as err:
                if debug: print('KeyError in remind/F', err)
            
            try:
                upload_row[columnToIndex('G')] = labeled_result['questions'][GC_INTRO]['questions'][QC_INTERVIEWER]['answer_label']
            except KeyError as err:
                if debug: print('KeyError in interviewer/G', err)
            
            try:
                upload_row[columnToIndex('N')] = labeled_result['questions'][GC_CONTACT]['questions'][QC_EMAIL]['answer_label']
            except KeyError as err:
                if debug: print('KeyError in email/N', err)
            
            try:
                upload_row[columnToIndex('O')] = labeled_result['questions'][GC_CONTACT]['questions'][QC_WHATSAPP]['answer_code']
            except KeyError as err:
                if debug: print('KeyError in whatsapp/O', err)
            
            try:
                upload_row[columnToIndex('P')] = labeled_result['questions'][GC_CONTACT]['questions'][QC_TELEGRAM]['answer_code']
            except KeyError as err:
                if debug: print('KeyError in telegram/P', err)
            
            try:
                upload_row[columnToIndex('Q')] = labeled_result['questions'][GC_CONTACT]['questions'][QC_FACEBOOK]['answer_code']
            except KeyError as err:
                if debug: print('KeyError in facebook/Q', err)
            
            try:
                upload_row[columnToIndex('R')] = labeled_result['questions'][GC_CONTACT]['questions'][QC_VIBER]['answer_code']
            except KeyError as err:
                if debug: print('KeyError in viber/R', err)
            
            try:
                upload_row[columnToIndex('S')] = labeled_result['questions'][GC_CONTACT]['questions'][QC_SMS]['answer_code']
            except KeyError as err:
                if debug: print('KeyError in sms/S', err)
            
            try:
                upload_row[columnToIndex('T')] = labeled_result['questions'][GC_CONTACT]['questions'][QC_CONTACT_OTHER]['answer_code']
            except KeyError as err:
                if debug: print('KeyError in other/T', err)
            
            try:
                upload_row[columnToIndex('U')] = labeled_result['questions'][GC_ABOUT]['questions'][QC_NATIONALITY]['answer_label']
            except KeyError as err:
                if debug: print('KeyError in nationality/U', err)
            
            request_body = {
                "range": 'List of repeats',
                "majorDimension": 'ROWS',
                'values': [upload_row]
            }
            request = sheet.values().append(spreadsheetId=GOOGLE_SHEET_IDS['CONTACTS'], range='List of repeats', valueInputOption='USER_ENTERED', insertDataOption='INSERT_ROWS', body=request_body)
            response = request.execute()
    
    db.execute(f"UPDATE lastrun SET lastsubmit = '{ labeled_results[-1]['meta']['_submission_time'] }'")
    conn.commit()
    conn.close()
    
    #print('3')
    #print('2')
    #print('1')
    pass