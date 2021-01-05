# Python functions to process response data

import koboextractor
import sys

from common import debugMsg, printMsgAndQuit

# Enable for verbose debug logging (disabled by default)
g_EnableDebugMsg = False

kobo = None


'''
Function to initialise the KoboExtractor.
'''
def init_kobo(token):
    global kobo
    kobo = koboextractor.KoboExtractor(token, 'https://kf.kobotoolbox.org/api/v2', debug=g_EnableDebugMsg)


'''
Return a list of assets as downloaded from Kobo.
'''
def get_assets(asset_uids):
    global kobo
    if kobo is None:
        printMsgAndQuit("Error in get_assets: KoboExtractor not initialised.")
    
    assets = []
    for asset_uid in asset_uids:
        assets.append(kobo.get_asset(asset_uid))
    return assets


'''
Function to move/rename a question in the data as downloaded from Kobo.
'''
def move_question(asset_data, from_code, to_code):
    for result in asset_data['results']:
        if from_code in result:
            # Copy question into new place
            result[to_code] = result[from_code]
            # Delete original
            del result[from_code]
    return asset_data


def get_asset_data(asset_uids):
    global kobo
    if kobo is None:
        printMsgAndQuit("Error in get_asset_data: KoboExtractor not initialised.")
    
    new_data = []
    for asset_uid in asset_uids:
        # TODO: Change back to this original line before deploying:
        #asset_data = kobo.get_data(asset_uid, submitted_after=last_submit_time)
        asset_data = kobo.get_data(asset_uid, query='{"_id":80353595}') # Get one of Heiko's responses for v34
        if asset_uid == 'aAYAW5qZHEcroKwvEq8pRb':
            # Special treatment to re-arrange the interview version
            asset_data = move_question(asset_data, 'S40/residence', 'S10/residence')
            asset_data = move_question(asset_data, 'S40/residence_99', 'S10/residence_99')
            asset_data = move_question(asset_data, 'S10/email', 'S70/email')
        new_data.append(asset_data)
    return new_data

'''
Function to count the number of new submissions in a list of Kobo asset data.
'''
def count_submissions(list_of_asset_data):
    new_submissions = 0
    for asset_data in list_of_asset_data:
        new_submissions += asset_data['count']
    return new_submissions


'''
Function to extract a merged list of all choices from all assets.
'''
def get_choice_lists(assets):
    global kobo
    if kobo is None:
        printMsgAndQuit("Error in get_choice_lists: KoboExtractor not initialised.")
    
    def merge_dicts(d1, d2):
        """
        Modifies d1 in-place to contain values from d2.  If any value
        in d1 is a dictionary (or dict-like), *and* the corresponding
        value in d2 is also a dictionary, then merge them in-place.
        """
        for k,v2 in d2.items():
            v1 = d1.get(k) # returns None if v1 has no value for this key
            if ( isinstance(v1, dict) and
                isinstance(v2, dict) ):
                merge_dicts(v1, v2)
            else:
                d1[k] = v2
    
    # Create dict of of choice options in the form of choice_lists[list_name][answer_code] = answer_label
    # Merge the choice lists into one, with the first versions taking precedence
    choice_lists = {}
    for asset in reversed(assets):
        merge_dicts(choice_lists, kobo.get_choices(asset))
    return choice_lists


def get_merged_questions(questions_list):
    def mergeQuestions(*questions_dicts):
        # Recursively merge the questions into one.
        # The sequence numbers will be colliding, hence the dicts cannot be
        # simply merged.
        # The first list of questions takes precedence; only additional
        # questions from the other lists will be added.
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

    questions = mergeQuestions(*questions_list)
    return questions


def get_questions_list(assets):
    global kobo
    if kobo is None:
        printMsgAndQuit("Error in get_merged_questions: KoboExtractor not initialised.")
    
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
    return questions_list
    

def get_labeled_result(new_data, choice_lists, questions, questions_list):
    global kobo
    if kobo is None:
        printMsgAndQuit("Error in get_choice_lists: KoboExtractor not initialised.")
    
    labeled_results = []
    for i in range(0, len(new_data)):
        for result in new_data[i]['results']:
            labeled_results.append(kobo.label_result(unlabeled_result=result, choice_lists=choice_lists, questions=questions_list[i], unpack_multiples=True))
    return sorted(labeled_results, key=lambda result: result['meta']['_submission_time'])
