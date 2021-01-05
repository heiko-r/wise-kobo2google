# Python script to prepare the tabular content for Google Sheets.

from common import GROUP_CODES, QUESTION_CODES, INSERT_UNIQUE_ID_AFTER, VERSIONS, FORMULAS, getUniqueId, debugMsg, printMsgAndQuit

from datetime import datetime, timedelta
import string

'''
Function to combine all questions and choices from all groups into one list.
'''
def flatten_questions(group_list):
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

'''
Function to get the code of a question or choice for use in the sheets header.
'''
def get_question_header_code(question):
    if 'type' in question:
        # question is a question
        new_code = '/'.join([question['group_code'], question['code']])
    else:
        # question is a choice
        new_code = '/'.join([question["group_code"], question['question_code'], question["code"]])
    return new_code

'''
Function to get the label of a question or choice for use in the sheets header.
'''
def get_question_header_label(question):
    if 'label' in question:
        new_label = question['label']
    else:
        new_label = question['code']
    return new_label

'''
Function to get the header values for the Cleaned Data sheets.
'''
def get_data_headers(flat_questions):
    sorted_codes = [None, None, None, None, None, None]
    sorted_labels = ['_submission_time', 'Submission date and time (computed)', 'Date (computed)', '_version_', 'Version (computed)', '_id']
    for question in flat_questions:
        new_code = get_question_header_code(question)
        new_label = get_question_header_label(question)
        if not question['group_code'].startswith(GROUP_CODES['contact']):
            sorted_codes.append(new_code)
            sorted_labels.append(new_label)
        # Add the unique ID and row # of last response
        if question['group_code'] == INSERT_UNIQUE_ID_AFTER['group'] and question['code'] == INSERT_UNIQUE_ID_AFTER['question']:
            sorted_codes.append('')
            sorted_labels.append('Unique ID (computed)')
            sorted_codes.append('')
            sorted_labels.append('Row # of latest response of respondent')
    sorted_values = [
        sorted_codes,
        sorted_labels
    ]
    return sorted_values

'''
Function to get the header values for the Respondent Contact Data sheets.
'''
def get_contact_headers(flat_questions):
    contact_codes = [None, None, None, None, None]
    contact_labels = ['_submission_time', 'Submission date and time (computed)', '_id', 'Unique ID (computed)', 'Interviewer']
    for question in flat_questions:
        new_code = get_question_header_code(question)
        new_label = get_question_header_label(question)
        if question['group_code'].startswith(GROUP_CODES['contact']):
            contact_codes.append(new_code)
            contact_labels.append(new_label)
    contact_values = [
        contact_codes,
        contact_labels
    ]
    return contact_values

'''
Function to get the survey version number from its code.
'''
def get_version(labeled_result):
    if '_version_' in labeled_result['meta'] and labeled_result['meta']['_version_'] in VERSIONS:
        version = VERSIONS[labeled_result['meta']['_version_']]
        version_ = labeled_result['meta']['_version_']
    elif '__version__' in labeled_result['meta'] and labeled_result['meta']['__version__'] in VERSIONS:
        version = VERSIONS[labeled_result['meta']['__version__']]
        version_ = labeled_result['meta']['__version__']
    else:
        version = 'NEW'
        version_ = 'N/A'
    return version, version_

'''
Function to get the label or code for an answer (question or choice).
'''
def get_answer(question, labeled_result, value_type):
    if 'type' in question:
        # question is a question, not a choice
        results_code = '/'.join([question['group_code'], question['code']])
        if results_code in labeled_result['results']:
            if value_type == "labels":
                new_value = labeled_result['results'][results_code]['answer_label']
            else:
                new_value = labeled_result['results'][results_code]['answer_code']
        else:
            new_value = ''
    else:
        # question is a choice
        results_code = '/'.join([question['group_code'], question['question_code']])
        if results_code in labeled_result['results'] and question['code'] in labeled_result['results'][results_code]['choices']:
            if value_type == "codes":
                new_value = labeled_result['results'][results_code]['choices'][question['code']]['answer_label']
            else:
                new_value = labeled_result['results'][results_code]['choices'][question['code']]['answer_code']
        else:
            new_value = ''
    return new_value

'''
Function to get the answer labels or codes for the Cleaned Data sheets.
'''
def get_data_values(labeled_results, flat_questions, value_type):
    if value_type != "labels" and value_type != "codes":
        printMsgAndQuit("Invalid value_type for get_data_values().")
    
    upload_values = []
    for labeled_result in labeled_results:
        # Determine version
        version, version_ = get_version(labeled_result)
        
        result_values = [labeled_result['meta']['_submission_time'], FORMULAS['datetime'], FORMULAS['date'], version_, version, labeled_result['meta']['_id']]
        
        for question in flat_questions:
            new_values = []
            new_values.append(get_answer(question, labeled_result, value_type))
            if not question['group_code'].startswith(GROUP_CODES['contact']):
                result_values += new_values
            # Add the unique ID and row # of last response
            if question['group_code'] == INSERT_UNIQUE_ID_AFTER['group'] and question['code'] == INSERT_UNIQUE_ID_AFTER['question']:
                result_values.append(getUniqueId(labeled_result))
                result_values.append(FORMULAS['latest_response'])
        upload_values.append(result_values)
    return upload_values

'''
Function to get the answer labels or codes for the Respondent Contact Data sheets.
'''
def get_contact_values(labeled_results, flat_questions, value_type):
    if value_type != "labels" and value_type != "codes":
        printMsgAndQuit("Invalid value_type for get_data_values().")
    
    contact_upload_values = []
    for labeled_result in labeled_results:
        if '/'.join([GROUP_CODES['intro'], QUESTION_CODES['interviewer']]) in labeled_result['results']:
            interviewer = labeled_result['results']['/'.join([GROUP_CODES['intro'], QUESTION_CODES['interviewer']])]['answer_code']
        else:
            interviewer = ''
        
        contact_values = [labeled_result['meta']['_submission_time'], FORMULAS['datetime'], labeled_result['meta']['_id'], getUniqueId(labeled_result), interviewer]
        
        for question in flat_questions:
            new_values = []
            new_values.append(get_answer(question, labeled_result, value_type))
            if question['group_code'].startswith(GROUP_CODES['contact']):
                contact_values += new_values
        contact_upload_values.append(contact_values)
    return contact_upload_values

'''
Function to get the numerical index of a sheet column from its A-Notation.
'''
def columnToIndex(column_string):
    index = -1
    for i in range(0, len(column_string)): # 0,1,2
        index += (string.ascii_uppercase.index(column_string[-i-1]) + 1) * (26**i)
    return index

'''
Function to get the values to add to the partners sheet.
'''
def get_partner_values(labeled_results):
    upload_values = []
    for labeled_result in labeled_results:
        if '/'.join([GROUP_CODES['contact'], QUESTION_CODES['partner']]) in labeled_result['results'] and labeled_result['results']['/'.join([GROUP_CODES['contact'], QUESTION_CODES['partner']])]['answer_code'] == '01':
            debugMsg('Found new partner!')
            # Person wants to partner -> add to the partner list
            upload_row = [None] * (columnToIndex('K') + 1) # Data list for columns A to K

            timestring = labeled_result['meta']['_submission_time']
            timeobj = datetime.fromisoformat(timestring)
            upload_row[columnToIndex('A')] = (timeobj + timedelta(hours=8)).date().isoformat()

            upload_row[columnToIndex('B')] = getUniqueId(labeled_result)

            upload_row[columnToIndex('C')] = labeled_result['results']['/'.join([GROUP_CODES['contact'], QUESTION_CODES['partner']])]['answer_label']

            sheet_columns = [
                (GROUP_CODES['contact'], QUESTION_CODES['orgname']),
                (GROUP_CODES['contact'], QUESTION_CODES['email']),
                (GROUP_CODES['contact'], QUESTION_CODES['whatsapp']),
                (GROUP_CODES['contact'], QUESTION_CODES['telegram']),
                (GROUP_CODES['contact'], QUESTION_CODES['facebook']),
                (GROUP_CODES['contact'], QUESTION_CODES['viber']),
                (GROUP_CODES['contact'], QUESTION_CODES['sms']),
                (GROUP_CODES['contact'], QUESTION_CODES['contact_other']),
            ]
            column_index = columnToIndex('D')
            for group, question in sheet_columns:
                try:
                    upload_row[column_index] = labeled_result['results']['/'.join([group, question])]['answer_label']
                except KeyError as err:
                    debugMsg(f'KeyError in partner sheet, question {question}:', err)
                column_index += 1

            upload_values.append(upload_row)
    return upload_values

'''
Function to get the values to add to the repeat respondents sheet.
'''
def get_repeat_values(labeled_results, uidlist):
    upload_values = []
    for labeled_result in labeled_results:
        if [getUniqueId(labeled_result)] in uidlist or ('/'.join([GROUP_CODES['intro'], QUESTION_CODES['before']]) in labeled_result['results'] and labeled_result['results']['/'.join([GROUP_CODES['intro'], QUESTION_CODES['before']])]['answer_code'] == '01') or ('/'.join([GROUP_CODES['contact'], QUESTION_CODES['again']]) in labeled_result['results'] and labeled_result['results']['/'.join([GROUP_CODES['contact'], QUESTION_CODES['again']])]['answer_code'] != '00'):
            # Respondent has answered before, claims to have answered before or wants to do the survey again
            debugMsg('Found (potential) repeat respondent!')
            # Add to repeat respondent list
            upload_row = [None] * (columnToIndex('N') + 1) # Data list for columns A to N

            timestring = labeled_result['meta']['_submission_time']
            timeobj = datetime.fromisoformat(timestring)
            day_of_submission = (timeobj + timedelta(hours=8)).date()
            upload_row[columnToIndex('A')] = day_of_submission.isoformat()

            upload_row[columnToIndex('B')] = getUniqueId(labeled_result)

            sheet_columns = [
                (GROUP_CODES['intro'], QUESTION_CODES['before']),
                (GROUP_CODES['contact'], QUESTION_CODES['again']),
                (GROUP_CODES['contact'], QUESTION_CODES['remind']),
                (GROUP_CODES['intro'], QUESTION_CODES['interviewer']),
                (GROUP_CODES['contact'], QUESTION_CODES['email']),
                (GROUP_CODES['contact'], QUESTION_CODES['whatsapp']),
                (GROUP_CODES['contact'], QUESTION_CODES['telegram']),
                (GROUP_CODES['contact'], QUESTION_CODES['facebook']),
                (GROUP_CODES['contact'], QUESTION_CODES['viber']),
                (GROUP_CODES['contact'], QUESTION_CODES['sms']),
                (GROUP_CODES['contact'], QUESTION_CODES['contact_other']),
                (GROUP_CODES['about'], QUESTION_CODES['nationality']),
            ]
            column_index = columnToIndex('C')
            for group, question in sheet_columns:
                try:
                    upload_row[column_index] = labeled_result['results']['/'.join([group, question])]['answer_label']
                except KeyError as err:
                    debugMsg(f'KeyError in repeats sheet, question {question}: ' + str(err))
                column_index += 1
            upload_values.append(upload_row)
    return upload_values
