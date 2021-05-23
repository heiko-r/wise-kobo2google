# Python file containing group and question codes used throughout the scripts.

import json
import os
import requests
import sys

from settings import *

# Global Constants
GOOGLE_SHEET_IDS_FILE_NAME = 'test-google-sheet-ids.json' # TODO: Change this back before deploying
GOOGLE_TOKENS_FILE_NAME = 'google_tokens.json'
MAIL_TEMPLATE_STYLES_FILE_NAME = 'mailtemplate-styles.json'
SMTP_CREDENTIALS_FILE_NAME = 'smtp-credentials.json'

GROUP_CODES = {
    'intro': 'S10',
    'about': 'S40',
    'id': 'S60',
    'contact': 'S70',
}
QUESTION_CODES = {
    'volunteer': 'volunteer',
    'partner': 'partner',
    'orgname': 'partner_org',
    'email': 'email',
    'whatsapp': 'whatsapp',
    'telegram': 'telegram',
    'facebook': 'fb',
    'viber': 'viber',
    'sms': 'sms',
    'contact_other': 'other',
    'copy': 'copy',
    'again': 'again',
    'remind': 'remind',
    'id1': 'ID1',
    'id2': 'ID2',
    'id3': 'ID3',
    'id4': 'ID4',
    'id5': 'ID5',
    'age': 'age',
    'nationality': 'nationality',
    'before': 'S10Q05',
    'interviewer': 'interviewer',
}

INSERT_UNIQUE_ID_AFTER = {
    'group': 'S60',
    'question': 'ID5'
}

VERSIONS = {
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
    'v7VzwKu2WZ7j9vEdTE6PoK': 'v6-dsm',
    'v3nJYwed3rK2tUydZyavQk': 'v35',
    'vGcHHKCPxRvbXXmTsyZVy9': 'v11-twc2',
    'vtBqADcuAAXVZsUg6FDm53': 'v11-i',
    'vAU5aZtEiGxGwr8q5dEmhW': 'v12-i',
    'vEP375D6umfmsv2oQtWmL7': 'v12-twc2',
    'vWP2hvkdP9ZqjpwZpror5K': 'v36',
    'vwfEcRUUwEfqd3Gjwrn9zv': 'v37',
    'vRr6gaC5NTmeSSYaMpL28B': 'v38',
    'vDNJjhcPSh4USSDtzL7hFB': 'v13-twc2',
    'vC8Gpx3iFvqEuv3sy88iEE': 'v13-i',
    'vnLMjcyhr6ywCz8fT4daEk': 'v13-twc2',
    'voAS4WQZmKeNEM89yR5k4p': 'v39',
    'v6KQv5mXYvnySw6UfPEX4q': 'v40',
    'vCnnME3ZqXKoXUrN2ihEjH': 'v14-twc2',
    'vj5i7u9LgXVHgJTtEfgmCk': 'v14-i',
}

COLUMNS = {
    'unique_id': 'JZ',
    'latest_response': 'KA',
}

FORMULAS = {
    'datetime': '=(TEXT(LEFT(INDIRECT("A"&ROW()),10),"yyyy-mm-dd ")&TEXT(RIGHT(INDIRECT("A"&ROW()),8),"hh:mm:ss"))+8/24',
    'date': '=(TEXT(LEFT(INDIRECT("B"&ROW()),10),"yyyy-mm-dd "))',
    'latest_response': f'=if(INDIRECT({COLUMNS["latest_response"]}$1&ROW())="","",max(arrayformula(if({COLUMNS["unique_id"]}:{COLUMNS["unique_id"]}=INDIRECT({COLUMNS["latest_response"]}$1&ROW()),row({COLUMNS["unique_id"]}:{COLUMNS["unique_id"]})))))',
}

'''
Function to get the Unique ID for a response.
'''
def getUniqueId(labeled_result):
    global GROUP_CODES, QUESTION_CODES
    debugMsg('Trying to extract ID')
    try:
        id1 = str(labeled_result['results']['/'.join([GROUP_CODES['id'], QUESTION_CODES['id1']])]['answer_code']).upper()
        id2 = "{0:0=2d}".format(int(labeled_result['results']['/'.join([GROUP_CODES['id'], QUESTION_CODES['id2']])]['answer_code']))
        id3 = str(labeled_result['results']['/'.join([GROUP_CODES['id'], QUESTION_CODES['id3']])]['answer_code'])
        id4 = str(labeled_result['results']['/'.join([GROUP_CODES['id'], QUESTION_CODES['id4']])]['answer_code'])
        if '/'.join([GROUP_CODES['id'], QUESTION_CODES['id5']]) in labeled_result['results']:
            id5 = str(labeled_result['results']['/'.join([GROUP_CODES['id'], QUESTION_CODES['id5']])]['answer_code'])
        else:
            id5 = str(labeled_result['results']['/'.join([GROUP_CODES['about'], QUESTION_CODES['age']])]['answer_code'][-1])
        debugMsg(f'ID: { id1 + id2 + id3 + id4 + id5 }')
        uniqueId = id1 + id2 + id3 + id4 + id5
    except KeyError as err:
        debugMsg('KeyError in Unique ID', err)
        uniqueId = ''
    return uniqueId


'''
Function to check internet connection.
Return True if connected to internet, False otherwise.
'''
def isConnectedToInternet(url='http://www.google.com/', timeout=5):
    try:
        requests.get(url, timeout=timeout)
    except requests.ConnectionError:
        return False

    return True

'''
Checks file presence for input file name list.
Return True if all files exists, False otherwise.
'''
def isFileExists(fileNameList):
    for fileName in fileNameList:
        if not os.path.exists(fileName):
            return False

    return True

'''
Perform environment checks to ensure all required settings are present.
'''
def checkEnvironment(requiredFileNameList=[], isInternetCheckRequired=False):
    # Check all required files present.
    if not isFileExists(requiredFileNameList):
        printMsgAndQuit("Error: Required file doesn't exists!")

    # Check internet connection.
    if isInternetCheckRequired and not isConnectedToInternet():
        printMsgAndQuit("Error: No internet connection.")

'''
Loads Json content from input file and returns a json object.
'''
def loadJson(fileName):
    with open(fileName, 'r') as f:
        return json.load(f)

    print("\tWarning: Failed to load Json from file: %s" % fileName)
    return None

'''
Prints debug verbose message.
'''
def debugMsg(message, err=None):
    if g_EnableDebugMsg:
        if err:
            print(message + " (error:%s)" % str(err))
        else:
            print(message)

'''
Print message and quit.
'''
def printMsgAndQuit(message, errorCode=-1):
    print("")
    print(message)
    sys.stdout.flush()
    sys.exit(errorCode)

'''
Quit program with error code.
'''
def quitApp(errorCode=-1):
    sys.exit(errorCode)
