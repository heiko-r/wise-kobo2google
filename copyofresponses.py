# Python script to process a given response into HTML, TXT and PDF templates
# The locale en_SG.UTF-8 needs to be installed on the system!

import jinja2
import locale
from datetime import datetime, timedelta
import json

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
def get_next_date(labeled_result, codes_constants):
    submission_time_iso = labeled_result['meta']['_submission_time']
    weeks = int(labeled_result['results']['/'.join([codes_constants['GC_CONTACT'], codes_constants['QC_AGAIN']])]['answer_code'])
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
def get_html(labeled_result, questions, flat_questions, codes_constants):
    init_template_env()
    html_template = template_env.get_template('mailtemplate.html')

    with open('mailtemplate-styles.json', 'r') as jsonfile:
        styles = json.load(jsonfile)
    
    locale.setlocale(locale.LC_ALL, LOCALE)

    next_date = get_next_date(labeled_result, codes_constants)

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
def get_txt(labeled_result, questions, flat_questions, codes_constants):
    init_template_env()
    txt_template = template_env.get_template('mailtemplate.txt')

    locale.setlocale(locale.LC_ALL, LOCALE)

    next_date = get_next_date(labeled_result, codes_constants)

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
def get_pdfhtml(labeled_result, questions, flat_questions, codes_constants):
    init_template_env()
    pdf_template = template_env.get_template('pdftemplate.html')

    with open('mailtemplate-styles.json', 'r') as jsonfile:
        styles = json.load(jsonfile)
    
    locale.setlocale(locale.LC_ALL, LOCALE)

    next_date = get_next_date(labeled_result, codes_constants)

    template_data = {
        'questions': questions,
        'flat_questions': flat_questions,
        'labeled_result': labeled_result,
        'next_date': next_date,
        'styles': styles
    }

    return pdf_template.render(template_data)
