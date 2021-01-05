# Python script to process a given response into HTML, TXT and PDF templates
# The locale en_SG.UTF-8 needs to be installed on the system!

import jinja2
import locale
from datetime import datetime, timedelta
import json

from common import GROUP_CODES, QUESTION_CODES, debugMsg
import sendmail
import pdfcopy
import googledrive

# Enable for verbose debug logging (disabled by default)
g_EnableDebugMsg = False

LOCALE = 'en_SG.UTF-8'

template_env = None


'''
Filter for debugging templates. Callable by Jinja2 templates.
Usage e.g.: {{ questions.S70.questions.age.answer_code | templatedebug }}
'''
def templatedebug(text):
    print(text)
    return ''


'''
Function to calculate the date for the next submission.
'''
def get_next_date(labeled_result):
    submission_time_iso = labeled_result['meta']['_submission_time']
    weeks = int(labeled_result['results']['/'.join([GROUP_CODES['contact'], QUESTION_CODES['again']])]['answer_code'])
    return (datetime.fromisoformat(submission_time_iso) + timedelta(days=weeks*7, hours=8)).strftime('%A, %d %B')


'''
Function to get the label for a group, question or choice.
'''
def get_label(questions, group_code, question_code=None, choice_code=None):
    def label_or_code(item, code):
        # If the item does not have a label, return the code instead
        if 'label' in item:
            return item['label']
        else:
            return code
    
    for group in questions:
        if group['code'] == group_code:
            if not question_code:
                return label_or_code(group, group_code)
            else:
                for question in group['questions']:
                    if question['code'] == 'question_code':
                        if not choice_code:
                            return label_or_code(question, question_code)
                        else:
                            for choice in question['choices']:
                                if choice['code'] == choice_code:
                                    return label_or_code(choice, choice_code)


'''
Function to initialise the Jinja2 template environment.
'''
def init_template_env():
    global template_env
    global g_EnableDebugMsg

    if template_env is None:
        template_loader = jinja2.FileSystemLoader('./')
        template_env = jinja2.Environment(
            loader = template_loader,
            autoescape = jinja2.select_autoescape(['html', 'xml'])
        )
        template_env.filters['debug'] = g_EnableDebugMsg
        template_env.globals['getLabel'] = get_label


'''
Function to generate the HTML version of responses.
'''
def get_html(labeled_result, questions, flat_questions):
    init_template_env()
    html_template = template_env.get_template('mailtemplate.html')

    with open('mailtemplate-styles.json', 'r') as jsonfile:
        styles = json.load(jsonfile)
    
    locale.setlocale(locale.LC_ALL, LOCALE)

    next_date = get_next_date(labeled_result)

    template_data = {
        'questions': questions,
        'flat_questions': flat_questions,
        'labeled_result': labeled_result,
        'next_date': next_date,
        'styles': styles
    }

    return html_template.render(template_data)


'''
Function to generate the plain-text version of responses.
'''
def get_txt(labeled_result, questions, flat_questions):
    init_template_env()
    txt_template = template_env.get_template('mailtemplate.txt')

    locale.setlocale(locale.LC_ALL, LOCALE)

    next_date = get_next_date(labeled_result)

    template_data = {
        'questions': questions,
        'flat_questions': flat_questions,
        'labeled_result': labeled_result,
        'next_date': next_date
    }

    return txt_template.render(template_data)


'''
Function to generate the HTML for the PDF version of responses.
'''
def get_pdfhtml(labeled_result, questions, flat_questions):
    init_template_env()
    pdf_template = template_env.get_template('pdftemplate.html')

    with open('mailtemplate-styles.json', 'r') as jsonfile:
        styles = json.load(jsonfile)
    
    locale.setlocale(locale.LC_ALL, LOCALE)

    next_date = get_next_date(labeled_result)

    template_data = {
        'questions': questions,
        'flat_questions': flat_questions,
        'labeled_result': labeled_result,
        'next_date': next_date,
        'styles': styles
    }

    return pdf_template.render(template_data)


'''
Function to send the copy of responses by email.
'''
def send_copy_by_email(labeled_result, questions, flat_questions):
    sender_email = 'covidsgsurvey@washinseasia.org'
    receiver_email = labeled_result['results']['/'.join([GROUP_CODES['contact'], QUESTION_CODES['email']])]['answer_label']
    subject = "Your responses - Survey on COVID-19 behaviours in Singapore"

    html = get_html(labeled_result, questions, flat_questions)
    txt = get_txt(labeled_result, questions, flat_questions)

    debugMsg(f'Sending mail to { receiver_email }')

    with open('smtp-credentials.json', 'r') as smtp_credentials_file:
        smtp_credentials = json.load(smtp_credentials_file)
    server = 'smtp.emaillabs.net.pl'
    port = 465
    smtp_conn = sendmail.connect_smtp(server, port, sendmail.ENCRYPTION_TLS, smtp_credentials['user'], smtp_credentials['password'])
    recipients = sendmail.send_email(smtp_conn, sender_email, receiver_email, subject, txt, html)
    if recipients is not None and len(recipients) > 0:
        debugMsg('Mail sent!')
    sendmail.disconnect_smtp(smtp_conn)

'''
Function to save copy of responses as PDF in Google Drive.
'''
def save_copy_as_pdf(labeled_result, questions, flat_questions):
    debugMsg('Copy requested, but no email given!')
    pdfhtml = get_pdfhtml(labeled_result, questions, flat_questions)
    fh = pdfcopy.create_pdf(pdfhtml)
    filename = f'{ getUniqueId(labeled_result) }.pdf'
    drive_file = googledrive.upload_file(filename, fh)
    debugMsg(f'File ID on Google Drive: { drive_file.get("id") }')

'''
Function to handle sending or saving the copy of responses.
'''
def handle_copies(labeled_results, questions, flat_questions):
    for labeled_result in labeled_results:
        if '/'.join([GROUP_CODES['contact'], QUESTION_CODES['copy']]) in labeled_result['results'] and labeled_result['results']['/'.join([GROUP_CODES['contact'], QUESTION_CODES['copy']])]['answer_code'] == '01':
            debugMsg('Copy requested!')
            if '/'.join([GROUP_CODES['contact'], QUESTION_CODES['email']]) in labeled_result['results']:
                # Send email automatically
                send_copy_by_email(labeled_result, questions, flat_questions)
            else:
                # no email address -> create PDF and add respondent to the list of copies to be sent
                save_copy_as_pdf(labeled_result, questions, flat_questions)
