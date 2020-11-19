# Make sure that the locale 'en_SG.UTF-8' is installed!

import os
import sys
from koboextractor import KoboExtractor
import sqlite3
import time
#import pickle # for Google auth
import json
from googleapiclient.discovery import build
#from google_auth_oauthlib.flow import InstalledAppFlow
import google.oauth2.credentials
from google.auth.transport.requests import Request
from googleapiclient.http import MediaIoBaseUpload
import string
from datetime import datetime, timedelta
import jinja2
import locale
import pdfkit
from io import BytesIO

# check for inhibit file
if os.path.isfile('inhibit'):
    exit()

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
        id1 = str(labeled_result['results']['/'.join([GC_ID, QC_ID1])]['answer_code']).upper()
        id2 = "{0:0=2d}".format(int(labeled_result['results']['/'.join([GC_ID, QC_ID2])]['answer_code']))
        id3 = str(labeled_result['results']['/'.join([GC_ID, QC_ID3])]['answer_code'])
        id4 = str(labeled_result['results']['/'.join([GC_ID, QC_ID4])]['answer_code'])
        if '/'.join([GC_ID, QC_ID5]) in labeled_result['results']:
            id5 = str(labeled_result['results']['/'.join([GC_ID, QC_ID5])]['answer_code'])
        else:
            id5 = str(labeled_result['results']['/'.join([GC_ABOUT, QC_AGE])]['answer_code'][-1])
        if debug: print(f'ID: { id1 + id2 + id3 + id4 + id5 }')
        uniqueId = id1 + id2 + id3 + id4 + id5
    except KeyError as err:
        if debug: print('KeyError in Unique ID', err)
        uniqueId = ''
    return uniqueId

def nextDate(submission_time_iso, weeks):
    return (datetime.fromisoformat(submission_time_iso) + timedelta(days=weeks*7, hours=8)).strftime('%A, %d %B')

def flattenQuestions(group_list):
    # Put all questions and choices from all groups into one list
    flat_questions = []
    for group in group_list:
        for question in group['questions']:
            new_question = question
            new_question['group_code'] = group['code']
            flat_questions.append(new_question)
            if question['choices']:
                for choice in question['choices']:
                    new_choice = choice
                    new_choice['group_code'] = group['code']
                    new_choice['question_code'] = question['code']
                    flat_questions.append(new_choice)
    return flat_questions

def mergeQuestions(*questions_dicts):
    # merge the questions into one.
    # sequence numbers will be colliding, hence the dicts cannot be simply merged
    # The first list of questions takes precedence; only additional questions from
    # the other lists will be added
    group_list = []
    
    for questions_dict in questions_dicts:
        # put all groups in a list:
        tmp_groups = []
        for group_code, group_dict in questions_dict['groups'].items():
            tmp_groups.append({
                'code': group_code,
                'sequence': group_dict['sequence'],
                'label': group_dict['label'],
                'repeat': group_dict['repeat']
            })
        # sort the list by sequence number:
        sorted_groups = sorted(tmp_groups, key=lambda group: group['sequence'])
        
        # go through sorted groups and add new ones to group_list:
        for new_group in sorted_groups:
            # search group_list for a group with the same code
            found_group = False
            for existing_group in group_list:
                if existing_group['code'] == new_group['code']:
                    found_group = True
                    break
            if not found_group:
                group_list.append({
                    'code': new_group['code'],
                    'label': new_group['label'],
                    'repeat': new_group['repeat'],
                    'questions': []
                })
        
        # go through each group and add new questions:
        for group in group_list:
            if group['code'] in questions_dict['groups'] and 'questions' in questions_dict['groups'][group['code']]:
                # put all questions in a list:
                tmp_questions = []
                for question_code, question_dict in questions_dict['groups'][group['code']]['questions'].items():
                    if group['code'] == 'S80' and question_code == 'logo': continue # skip strange 'logo' note at the end
                    tmp_question = {
                        'code': question_code,
                        'sequence': question_dict['sequence']
                    }
                    if 'label' in question_dict:
                        tmp_question['label'] = question_dict['label']
                    if 'type' in question_dict:
                        tmp_question['type'] = question_dict['type']
                    tmp_questions.append(tmp_question)
                # sort the list by sequence number:
                sorted_questions = sorted(tmp_questions, key=lambda question: question['sequence'])
                
                # go through sorted questions and add new ones to the group's list of questions:
                for tmp_question in sorted_questions:
                    # search existing question list for a question with the same code
                    found_question = False
                    for existing_question in group['questions']:
                        if existing_question['code'] == tmp_question['code']:
                            found_question = True
                            break
                    if not found_question:
                        new_question = {
                            'code': tmp_question['code'],
                            'choices': []
                        }
                        if 'label' in tmp_question:
                            new_question['label'] = tmp_question['label']
                        if 'type' in tmp_question:
                            new_question['type'] = tmp_question['type']
                        group['questions'].append(new_question)
        
        # go through each question and add new choices:
        for group in group_list:
            if group['code'] in questions_dict['groups'] and 'questions' in questions_dict['groups'][group['code']]:
                for question in group['questions']:
                    if question['code'] in questions_dict['groups'][group['code']]['questions'] and 'choices' in questions_dict['groups'][group['code']]['questions'][question['code']]:
                        # put all choices in a list:
                        tmp_choices = []
                        for choice_code, choice_dict in questions_dict['groups'][group['code']]['questions'][question['code']]['choices'].items():
                            tmp_choices.append({
                                'code': choice_code,
                                'label': choice_dict['label'],
                                'sequence': choice_dict['sequence']
                            })
                        # sort the list by sequence number:
                        sorted_choices = sorted(tmp_choices, key=lambda choice: choice['sequence'])
                        
                        # go through sorted choices and add new ones to the question's list of choices:
                        for new_choice in sorted_choices:
                            # search existin list for a choice with the same code
                            found_choice = False
                            for existing_choice in question['choices']:
                                if existing_choice['code'] == new_choice['code']:
                                    found_choice = True
                                    break
                            if not found_choice:
                                question['choices'].append({
                                    'code': new_choice['code'],
                                    'label': new_choice['label']
                                })
        # TODO: treat 'or other' somehow
    
    return group_list

def mergeDicts(d1, d2):
    """
    Modifies d1 in-place to contain values from d2.  If any value
    in d1 is a dictionary (or dict-like), *and* the corresponding
    value in d2 is also a dictionary, then merge them in-place.
    """
    for k,v2 in d2.items():
        v1 = d1.get(k) # returns None if v1 has no value for this key
        if ( isinstance(v1, dict) and 
             isinstance(v2, dict) ):
            mergeDicts(v1, v2)
        else:
            d1[k] = v2

def getLabel(merged_questions, group_code, question_code=None, choice_code=None):
    def labelOrCode(item, code):
        if 'label' in item:
            return item['label']
        else:
            return code
    
    for group in merged_questions:
        if group['code'] == group_code:
            if not question_code:
                return labelOrCode(group, group_code)
            else:
                for question in group['questions']:
                    if question['code'] == 'question_code':
                        if not choice_code:
                            return labelOrCode(question, question_code)
                        else:
                            for choice in question['choices']:
                                if choice['code'] == choice_code:
                                    return labelOrCode(choice, choice_code)
    
try:
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
    
    asset_uids = [
        'argHw9ZzcAtcmEytJbWQo7', # online version
        'aAYAW5qZHEcroKwvEq8pRb', # interview version
        'aKAexXkCDUM5WbL3jQW2V5', # DSM version
        'aXbWBpzEm8xZyatgciGnEd', # TWC2 version
    ]
    
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
    new_data = []
    for asset_uid in asset_uids:
        asset_data = kobo.get_data(asset_uid, submitted_after=last_submit_time)
        if asset_uid == 'aAYAW5qZHEcroKwvEq8pRb':
            # Special treatment for interview version
            # Move S40/residence, S40/residence_99 to S10/residence, S10/residence_99
            for result in asset_data['results']:
                if 'S40/residence' in result:
                    print("Found S40/residence")
                    result['S10/residence'] = result['S40/residence']
                    del result['S40/residence']
                if 'S40/residence_99' in result:
                    print("Found S40/residence_99")
                    result['S10/residence_99'] = result['S40/residence_99']
                    del result['S40/residence_99']
                # Move S10/email to S70/email
                if 'S10/email' in result:
                    print("Found S10/email")
                    result['S70/email'] = result['S10/email']
                    del result['S10/email']
        new_data.append(asset_data)
    #new_data_online = kobo.get_data(asset_uid_online, query='{"_submission_time": {"$gt": "2020-06-08T05:40:54", "$lt": "2020-06-08T06:11:09"}}')
    
    new_submissions = 0
    for new_data_results in new_data:
        new_submissions += new_data_results['count']
    
    if new_submissions > 0 or add_headers:
        if debug: print('Kobo: Getting assets')
        assets = []
        for asset_uid in asset_uids:
            assets.append(kobo.get_asset(asset_uid))
        
        if debug: print('Kobo: Getting labels for questions, choices and answers')
        
        ######## CHOICES ########
        # Create dict of of choice options in the form of choice_lists[list_name][answer_code] = answer_label
        # Merge the choice lists into one, with the first versions taking precedence
        choice_lists = {}
        for asset in reversed(assets):
            mergeDicts(choice_lists, kobo.get_choices(asset))
        if debug:
            print('choice_lists:')
            pprint.pprint(choice_lists)
        
        ######## QUESTIONS ########
        questions_list = []
        for asset in assets:
            asset_questions = kobo.get_questions(asset=asset, unpack_multiples=True)
            if asset['uid'] == 'aAYAW5qZHEcroKwvEq8pRb':
                # Special treatment for interview version
                # Move S40/residence, S40/residence_99 to S10/residence, S10/residence_99
                # Move S10/email to S70/email
                asset_questions['groups']['S10']['questions']['residence'] = asset_questions['groups']['S40']['questions']['residence']
                asset_questions['groups']['S10']['questions']['residence_99'] = asset_questions['groups']['S40']['questions']['residence_99']
                asset_questions['groups']['S70']['questions']['email'] = asset_questions['groups']['S10']['questions']['email']
                del asset_questions['groups']['S40']['questions']['residence']
                del asset_questions['groups']['S40']['questions']['residence_99']
                del asset_questions['groups']['S10']['questions']['email']
            questions_list.append(asset_questions)
        questions = mergeQuestions(*questions_list)
        
        ## Remove all questions without labels or of the following types
        ## Note: Types in use to keep 'note', 'select_one', 'select_multiple', 'integer', 'text', '_or_other'
        #delete_types = ['start', 'end', 'today', 'begin_group', 'end_group', 'calculate']
        #for question_group, question_group_dict in questions.items():
        #    # The [] part is building a list of question_codes where the question type is in the above delete list
        #    for question_code in [question_code for question_code, question_dict in question_group_dict['questions'].items() if question_dict['type'] in delete_types]: del questions[question_group]['questions'][question_code]
        #    for question_code in [question_code for question_code, question_dict in question_group_dict['questions'].items() if 'label' not in question_dict]: del questions[question_group]['questions'][question_code]
        ## delete empty question groups
        #for question_group in [question_group for question_group, question_group_dict in questions.items() if not question_group_dict['questions']]: del questions[question_group]
        if debug:
            print('Merged questions:')
            pprint.pprint(questions)
        
        flat_questions = flattenQuestions(questions)
        
        ######## INITIALISE GOOGLE ########
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
            'vwDorFQsYqPkrPqTpPSwA7': 'v3-i',
            'vVbU7nhpFDTRrHdTWWhSzU': 'v23',
            'vQBQZaxKXmjfrd6EUv2AyW': 'v24',
            'vuwzFJtrbydb7He29syKB7': 'v25',
            'vDRjcUWz5TfrmsHVETJaxL': 'v26',
            'vYkEdFB2Y3PRzbPGxo7GxA': 'v27',
            'v6kNJDKrrcEaxeBC6q4yyX': 'v4-i',
            'vAtyPtYBTyGBjgCHtU6nnK': 'v28',
            'v6TWK7RVicgPemLR6QzdQ4': 'v5-i',
            'vx9VE7g4ZdkiPt9S5G8HgR': 'v29',
            'vTLWtFpJETVBLGyY64y5iM': 'v6-i',
            'vfAt6ncFCDobgnEmVE532p': 'v30',
            'vHTNob6PmRp47CagSqrePt': 'v31',
            'vZ3kcYASkiWzdZbUzq3Xf9': 'v2-dsm',
            'vpKMUDvHe2BDri6MuqeL6s': 'v7-i',
            'vm6oMVrEYvLm833Z3UmHWT': 'v3-twc2',
            'vjqirEyxbRaEeSwaMWqwLL': 'v4-twc2',
            'vcPgrbRYNhLvZD3wEQQkd3': 'v5-twc2',
            'vjYKbRn8YkfoEi94DZDU97': 'v6-twc2',
            'vJPAyGJKHGNXefNie8Nuaq': 'v7-twc2',
            'vwcAGGT4juWtYrECiRf7XU': 'v8-twc2',
            'vMjX3eAnsvNNFVEXLNTQZD': 'v32',
            'vC9GgrbXQp6fVd2C4PGAt4': 'v3-dsm',
            'vrJXmnGFxRLEp2CL4bMQqF': 'v8-i',
            'vzpv5jeqgRpX6jDzBqxUVY': 'v33',
            'vnxPqQeWqgDbMkEKPG6Qo6': 'v9-twc2',
            'voYSHveEczpYBdWEJqKWvc': 'v4-dsm',
            'veGqP9bqcGBwRyRbwthwvX': 'v9-i',
            'viCb3faKAu4UQuiokmoQyZ': 'v10-i',
            'v5eQCAnx6QNRvvrYX89whb': 'v10-twc2',
            'vfsJCoq3dS4B5iHXHh4k4M': 'v5-dsm',
            'vYWjoWYNYymVsZ46EF5rSU': 'v34',
        }
        GOOGLE_UNIQUEID_AFTER_GROUP = 'S60'
        GOOGLE_UNIQUEID_AFTER_QUESTION = 'ID5'
        
        GOOGLE_COLUMN_UNIQUEID = 'JZ'
        GOOGLE_COLUMN_LATESTRESPONSE = 'KA'
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
        
        # Set up Google API authentication
        with open('google_tokens.json', 'r') as tokenfile:
            google_tokens = json.load(tokenfile)
        creds = google.oauth2.credentials.Credentials(
            None,
            refresh_token=google_tokens['refresh_token'],
            token_uri=google_tokens['token_uri'],
            client_id=google_tokens['client_id'],
            client_secret=google_tokens['client_secret']
        )
        
        # Create the Google Sheets API
        sheets_service = build('sheets', 'v4', credentials=creds)
        
        # Call the Sheets API
        sheet = sheets_service.spreadsheets()
        
        # Get list of existing UniqueIDs for later
        uidresult = sheet.values().get(spreadsheetId=GOOGLE_SHEET_IDS['CLEANEDDATA'], range=f'Data (labeled)!{GOOGLE_COLUMN_UNIQUEID}4:{GOOGLE_COLUMN_UNIQUEID}').execute()
        uidlist = uidresult.get('values', [])
        if debug:
            print("List of UIDs:")
            pprint.pprint(uidlist)
    
    
    if add_headers:
        # Sheet is still empty -> fill the header rows first
        # Insert two new rows on top to preserve whatever may be in the sheet already
        if debug: print('Adding headers')
        
        sorted_codes = [None, None, None, None, None, None]
        sorted_labels = ['_submission_time', 'Submission date and time (computed)', 'Date (computed)', '_version_', 'Version (computed)', '_id']
        contact_codes = [None, None, None, None, None]
        contact_labels = ['_submission_time', 'Submission date and time (computed)', '_id', 'Unique ID (computed)', 'Interviewer']
        for question in flat_questions:
            if 'type' in question:
                # question is a question
                new_code = '/'.join([question['group_code'], question['code']])
            else:
                # question is a choice
                new_code = '/'.join([question["group_code"], question['question_code'], question["code"]])
            if 'label' in question:
                new_label = question['label']
            else:
                new_label = question['code']
            if not question['group_code'].startswith(GC_CONTACT):
                sorted_codes.append(new_code)
                sorted_labels.append(new_label)
            else:
                contact_codes.append(new_code)
                contact_labels.append(new_label)
            # Add the unique ID and row # of last response
            if question['group_code'] == GOOGLE_UNIQUEID_AFTER_GROUP and question['code'] == GOOGLE_UNIQUEID_AFTER_QUESTION:
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
        for sheet_range in ['Data (labeled)', 'Data (codes)']:
            request_body = {
                "range": sheet_range,
                "majorDimension": 'ROWS',
                'values': sorted_values
            }
            request = sheet.values().append(spreadsheetId=GOOGLE_SHEET_IDS['CLEANEDDATA'], range=sheet_range, valueInputOption='USER_ENTERED', insertDataOption='INSERT_ROWS', body=request_body)
            response = request.execute()
        for sheet_range in ['Data (labeled)', 'Data (codes)']:
            request_body = {
                "range": sheet_range,
                "majorDimension": 'ROWS',
                'values': contact_values
            }
            request = sheet.values().append(spreadsheetId=GOOGLE_SHEET_IDS['S70'], range=sheet_range, valueInputOption='USER_ENTERED', insertDataOption='INSERT_ROWS', body=request_body)
            response = request.execute()
        exit()
        
    if new_submissions > 0:
        if debug: print(f'Labeling { new_submissions } submissions')
        # Concatenate the results from both versions and sort by submission time
        # Add the labels for questions and choices
        labeled_results = []
        for i in range(0, len(new_data)):
            for result in new_data[i]['results']:
                labeled_results.append(kobo.label_result(unlabeled_result=result, choice_lists=choice_lists, questions=questions_list[i], unpack_multiples=True))
        labeled_results = sorted(labeled_results, key=lambda result: result['meta']['_submission_time'])
        
        # Unpack and upload results
        for sheet_range in ['Data (labeled)', 'Data (codes)']:
            if debug: print(f'Uploading {new_submissions} submissions to {sheet_range}')
            upload_values = []
            contact_upload_values = []
            for labeled_result in labeled_results:
                if '_version_' in labeled_result['meta'] and labeled_result['meta']['_version_'] in GOOGLE_VERSION_DICT:
                    version = GOOGLE_VERSION_DICT[labeled_result['meta']['_version_']]
                    version_ = labeled_result['meta']['_version_']
                elif '__version__' in labeled_result['meta'] and labeled_result['meta']['__version__'] in GOOGLE_VERSION_DICT:
                    version = GOOGLE_VERSION_DICT[labeled_result['meta']['__version__']]
                    version_ = labeled_result['meta']['__version__']
                else:
                    version = 'NEW'
                    version_ = 'N/A'
                result_values = [labeled_result['meta']['_submission_time'], GOOGLE_DATETIME_FORMULA, GOOGLE_DATE_FORMULA, version_, version, labeled_result['meta']['_id']]
                if '/'.join([GC_INTRO, QC_INTERVIEWER]) in labeled_result['results']:
                    interviewer = labeled_result['results']['/'.join([GC_INTRO, QC_INTERVIEWER])]['answer_code']
                else:
                    interviewer = ''
                contact_values = [labeled_result['meta']['_submission_time'], GOOGLE_DATETIME_FORMULA, labeled_result['meta']['_id'], getUniqueId(labeled_result), interviewer]
                for question in flat_questions:
                    new_values = []
                    if 'type' in question:
                        # question is a question, not a choice
                        results_code = '/'.join([question['group_code'], question['code']])
                        if results_code in labeled_result['results']:
                            if sheet_range == 'Data (labeled)':
                                new_values.append(labeled_result['results'][results_code]['answer_label'])
                            else:
                                new_values.append(labeled_result['results'][results_code]['answer_code'])
                        else:
                            new_values.append('')
                    else:
                        # question is a choice
                        results_code = '/'.join([question['group_code'], question['question_code']])
                        if results_code in labeled_result['results'] and question['code'] in labeled_result['results'][results_code]['choices']:
                            if sheet_range == 'Data (labeled)':
                                new_values.append(labeled_result['results'][results_code]['choices'][question['code']]['answer_label'])
                            else:
                                new_values.append(labeled_result['results'][results_code]['choices'][question['code']]['answer_code'])
                        else:
                            new_values.append('')
                    if not question['group_code'].startswith(GC_CONTACT):
                        result_values += new_values
                    else:
                        contact_values += new_values
                    # Add the unique ID and row # of last response
                    if question['group_code'] == GOOGLE_UNIQUEID_AFTER_GROUP and question['code'] == GOOGLE_UNIQUEID_AFTER_QUESTION:
                        result_values.append(getUniqueId(labeled_result))
                        result_values.append(GOOGLE_LATEST_RESPONSE_FORMULA)
                upload_values.append(result_values)
                contact_upload_values.append(contact_values)
            
            #print('!')
            #print('.')
            #print('!')
            
            request_body = {
                "range": sheet_range,
                "majorDimension": 'ROWS',
                'values': upload_values
            }
            request = sheet.values().append(spreadsheetId=GOOGLE_SHEET_IDS['CLEANEDDATA'], range=sheet_range, valueInputOption='USER_ENTERED', insertDataOption='INSERT_ROWS', body=request_body)
            response = request.execute()
            
            request_body = {
                "range": sheet_range,
                "majorDimension": 'ROWS',
                'values': contact_upload_values
            }
            request = sheet.values().append(spreadsheetId=GOOGLE_SHEET_IDS['S70'], range=sheet_range, valueInputOption='USER_ENTERED', insertDataOption='INSERT_ROWS', body=request_body)
            response = request.execute()
        
        # Update partner list
        upload_values = []
        for labeled_result in labeled_results:
            if '/'.join([GC_CONTACT, QC_PARTNER]) in labeled_result['results'] and labeled_result['results']['/'.join([GC_CONTACT, QC_PARTNER])]['answer_code'] == '01':
                if debug: print('Found new partner!')
                # Person wants to partner -> add to the partner list
                upload_row = [None] * (columnToIndex('K') + 1) # Data list for columns A to K
                
                timestring = labeled_result['meta']['_submission_time']
                timeobj = datetime.fromisoformat(timestring)
                upload_row[columnToIndex('A')] = (timeobj + timedelta(hours=8)).date().isoformat()
                
                upload_row[columnToIndex('B')] = getUniqueId(labeled_result)
                
                upload_row[columnToIndex('C')] = labeled_result['results']['/'.join([GC_CONTACT, QC_PARTNER])]['answer_label']
                
                sheet_columns = [
                    (GC_CONTACT, QC_ORGNAME),
                    (GC_CONTACT, QC_EMAIL),
                    (GC_CONTACT, QC_WHATSAPP),
                    (GC_CONTACT, QC_TELEGRAM),
                    (GC_CONTACT, QC_FACEBOOK),
                    (GC_CONTACT, QC_VIBER),
                    (GC_CONTACT, QC_SMS),
                    (GC_CONTACT, QC_CONTACT_OTHER),
                ]
                column_index = columnToIndex('D')
                for group, question in sheet_columns:
                    try:
                        upload_row[column_index] = labeled_result['results']['/'.join([group, question])]['answer_label']
                    except KeyError as err:
                        if debug: print(f'KeyError in partner sheet, question {question}:', err)
                    column_index += 1
                
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
            if '/'.join([GC_CONTACT, QC_COPY]) in labeled_result['results'] and labeled_result['results']['/'.join([GC_CONTACT, QC_COPY])]['answer_code'] == '01':
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
                template_env.globals['getLabel'] = getLabel
                html_template = template_env.get_template('mailtemplate.html')
                txt_template = template_env.get_template('mailtemplate.txt')
                pdf_template = template_env.get_template('pdftemplate.html')
                
                with open('mailtemplate-styles.json', 'r') as jsonfile:
                    styles = json.load(jsonfile)
                
                locale.setlocale(locale.LC_ALL, 'en_SG.UTF-8')
                
                next_date = nextDate(labeled_result['meta']['_submission_time'], int(labeled_result['results']['/'.join([GC_CONTACT, QC_AGAIN])]['answer_code']))
                
                template_data = {
                    'questions': questions,
                    'flat_questions': flat_questions,
                    'labeled_result': labeled_result,
                    'next_date': next_date,
                    'styles': styles
                }
                
                if '/'.join([GC_CONTACT, QC_EMAIL]) in labeled_result['results']:
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
                    receiver_email = labeled_result['results']['/'.join([GC_CONTACT, QC_EMAIL])]['answer_label']
                    #receiver_email = 'heiko@rothkranz.net'
                    '''Format of gmail-credentials.json:
                    {
                        "user": "user@example.com",
                        "password": "safEpa55w0rd"
                    }'''
                    with open('gmail-credentials.json', 'r') as gmail_credentials_file:
                        gmail_credentials = json.load(gmail_credentials_file)
                    server = 'smtp.gmail.com'
                    port = 465
                    #gmail_credentials = {
                    #    'user': '40784a442d78cc',
                    #    'password': '2c2ea9a8135d5c'
                    #}
                    #server = 'smtp.mailtrap.io'
                    
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
            
            ######## LIST OF REPEATS ########
            if [getUniqueId(labeled_result)] in uidlist or ('/'.join([GC_INTRO, QC_BEFORE]) in labeled_result['results'] and labeled_result['results']['/'.join([GC_INTRO, QC_BEFORE])]['answer_code'] == '01') or ('/'.join([GC_CONTACT, QC_AGAIN]) in labeled_result['results'] and labeled_result['results']['/'.join([GC_CONTACT, QC_AGAIN])]['answer_code'] != '00'):
                # Respondent has answered before, claims to have answered before or wants to do the survey again
                if debug: print('Found (potential) repeat respondent!')
                # Add to repeat respondent list
                upload_row = [None] * (columnToIndex('N') + 1) # Data list for columns A to N
                
                timestring = labeled_result['meta']['_submission_time']
                timeobj = datetime.fromisoformat(timestring)
                day_of_submission = (timeobj + timedelta(hours=8)).date()
                upload_row[columnToIndex('A')] = day_of_submission.isoformat()
                
                upload_row[columnToIndex('B')] = getUniqueId(labeled_result)
                
                sheet_columns = [
                    (GC_INTRO, QC_BEFORE),
                    (GC_CONTACT, QC_AGAIN),
                    (GC_CONTACT, QC_REMIND),
                    (GC_INTRO, QC_INTERVIEWER),
                    (GC_CONTACT, QC_EMAIL),
                    (GC_CONTACT, QC_WHATSAPP),
                    (GC_CONTACT, QC_TELEGRAM),
                    (GC_CONTACT, QC_FACEBOOK),
                    (GC_CONTACT, QC_VIBER),
                    (GC_CONTACT, QC_SMS),
                    (GC_CONTACT, QC_CONTACT_OTHER),
                    (GC_ABOUT, QC_NATIONALITY),
                ]
                column_index = columnToIndex('C')
                for group, question in sheet_columns:
                    try:
                        upload_row[column_index] = labeled_result['results']['/'.join([group, question])]['answer_label']
                    except KeyError as err:
                        if debug: print(f'KeyError in repeats sheet, question {question}:', err)
                    column_index += 1
                
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
except:
    # something went wrong -> disable the tool by creating a file named 'inhibit'
    open('inhibit', 'a').close()
    raise