# Python file containing group and question codes used throughout the scripts.

import sys

# Enable for verbose debug logging (disabled by default)
g_EnableDebugMsg = False

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
        id1 = str(labeled_result['results']['/'.join([GROUP_CODES['id'], QUESTION_CODES['ID1']])]['answer_code']).upper()
        id2 = "{0:0=2d}".format(int(labeled_result['results']['/'.join([GROUP_CODES['id'], QUESTION_CODES['ID2']])]['answer_code']))
        id3 = str(labeled_result['results']['/'.join([GROUP_CODES['id'], QUESTION_CODES['ID3']])]['answer_code'])
        id4 = str(labeled_result['results']['/'.join([GROUP_CODES['id'], QUESTION_CODES['ID4']])]['answer_code'])
        if '/'.join([GROUP_CODES['id'], QUESTION_CODES['ID5']]) in labeled_result['results']:
            id5 = str(labeled_result['results']['/'.join([GROUP_CODES['id'], QUESTION_CODES['ID5']])]['answer_code'])
        else:
            id5 = str(labeled_result['results']['/'.join([GROUP_CODES['about'], QUESTION_CODES['age']])]['answer_code'][-1])
        debugMsg(f'ID: { id1 + id2 + id3 + id4 + id5 }')
        uniqueId = id1 + id2 + id3 + id4 + id5
    except KeyError as err:
        debugMsg('KeyError in Unique ID', err)
        uniqueId = ''
    return uniqueId


'''
Prints debug verbose message.
'''
def debugMsg(message, err = None):
    if g_EnableDebugMsg:
        if err is None:
            print(message)
        else:
            print(message + " " + str(err))

'''
Print message and quit.
'''
def printMsgAndQuit(message, errorCode=-1):
    print("")
    print(message)
    sys.stdout.flush()
    sys.exit(errorCode)