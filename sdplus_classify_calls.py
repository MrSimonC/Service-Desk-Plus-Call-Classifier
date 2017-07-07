import datetime
import os
import pyperclip
import re
import subprocess
import sys
import tempfile
from collections import OrderedDict
from custom_modules.sdplus_api_rest import API
from custom_modules.xlsx import XlsxTools
from typing import Dict, List
__version__ = '1.0'


def classify_call(conversations: list):
    """
    classify this call
    :param conversations: list of dicts from sdplus_api.request_get_all_conversations
    :return: classification string
    """
    for no in range(len(conversations)-1, -1, -1):
        conversation = conversations[no]
        if conversation['from'] == 'System':  # miss any 'System' notifications
            continue
        if conversation['from'] != 'IT Third Party Response':  # i.e. processed incomming CSC email
            continue

        if 'We have resolved the following case:' in conversation['description']:
            return 'closed'
        elif 'has been rejected, Please resubmit the case' in conversation['description']:
            return 'rejected'
        else:
            return 'open'
    return 'No CSC entries'


def find_all_people_involved(conversations: list, requester: str):
    # Find who else is involved in the call other than CSC and the requester (e.g. a Back Office member)
    all_names = []
    for no in range(len(conversations)-1, -1, -1):
        conversation = conversations[no]
        if conversation['from'] == 'System' \
                or conversation['from'] == 'IT Third Party Response' \
                or conversation['from'] == 'CSSC NHS IT Helpdesk' \
                or conversation['from'] == 'CSC Lorenzo Upgrade Support Service':
            continue
        if conversation['from'] != requester:
            all_names.append(conversation['from'])
    return ', '.join(list(set(all_names)))


def find_date_csc_opened_call(conversations: list):  # or reopened
    for no in range(len(conversations)-1, -1, -1):
        conversation = conversations[no]
        if conversation['from'] == 'IT Third Party Response':
            if 'has been created from triage' in conversation['description'] \
                    or 'We have re-open the following case' in conversation['description']:
                return conversation['createddate']
    return None


def find_csc_severity(conversations: list):
    for no in range(len(conversations)-1, -1, -1):
        conversation = conversations[no]
        if conversation['from'] == 'IT Third Party Response':
            if 'has been created from triage' in conversation['description'] \
                    or 'We have re-open the following case' in conversation['description']:
                if re.search(r'(?:Severity \: )(\d{1})', conversation['description']):
                    return re.search(r'(?:Severity \: )(\d{1})', conversation['description']).group(1)
                else:
                    return ''
    return ''


def process_calls():
    """
    Go through all "Back Office Third Party/CSC" calls, classifying if open or closed
    :return:
    """
    try:
        sdplus_api = API(os.environ['SDPLUS_ADMIN'], 'http://sdplus/sdpapi/')
        if not sdplus_api:
            raise KeyError
    except KeyError:
        print('Windows environment varible for "SDPLUS_ADMIN" (the API key for sdplus) wasn\'t found. \n'
              'Please correct using ""setx SDPLUS_ADMIN <insert your own SDPLUS key here>" in a command line.')
        sys.exit(1)
    result = []
    all_queues = sdplus_api.request_get_requests('Back Office Third Party/CSC_QUEUE')
    for each_call in all_queues:
        conversations = sdplus_api.request_get_all_conversations(each_call['workorderid'])
        each_call['classification'] = classify_call(conversations)
        each_call['Others involved'] = find_all_people_involved(conversations, each_call['requester'])
        each_call['CSC open/reopen date'] = find_date_csc_opened_call(conversations)
        each_call['CSC severity'] = find_csc_severity(conversations)
        result.append(each_call)
    return result


def output_to_temp_xlsx_file(contents: List[Dict], timestamp=''):
    if not timestamp:
        timestamp = datetime.datetime.now().strftime('%d%b%y_%H%M')
    results_file = os.path.join(tempfile.gettempdir(), 'csc_calls_analysis_{0}.xlsx'.format(timestamp))
    xlsx = XlsxTools()
    xlsx.create_document(contents, 'csc analysis', results_file)
    return results_file


if __name__ == '__main__':
    print('SDPlus CSC Call Classifier v' + __version__ + '\n'
          'Taking to SDPlus, getting call information on Back Office Third Party/CSC queue')
    classified_csc_calls = process_calls()
    print('Processing data')
    header = ['Analysis Date',
              'Ref',
              'Classification',
              'Requester',
              'Others involved',
              'Raised date',
              'CSC severity',
              'CSC open/reopen date',
              'Work days since',
              'Subject']
    results = []
    timestamp = datetime.datetime.now().strftime('%d%b%y_%H%M')
    for number, each in enumerate(classified_csc_calls):
        network_days_column = 'H'
        network_days_number = number + 2
        results.append(['(Report run: ' + timestamp + ')',
                       each['workorderid'],
                       each['classification'],
                       each['requester'],
                       each['Others involved'],
                       each['createdtime'].strftime('%d/%b/%Y'),
                       each['CSC severity'],
                       each['CSC open/reopen date'].strftime('%d/%b/%Y') if each['CSC open/reopen date'] else '',
                       '=IF(NOT(ISBLANK({col}{row})), NETWORKDAYS({col}{row},NOW()), "")'.format(
                           col=network_days_column,
                           row=network_days_number),
                       each['subject']])
    text_output = '\t'.join(header) + '\n'
    excel_contents = []
    for each_row in results:
        text_output += '\t'.join(each_row) + '\n'
        excel_contents.append(OrderedDict(zip(header, each_row)))
    print(text_output)
    pyperclip.copy(text_output)
    print('Results copied to clipboard.')
    excel_results_file = output_to_temp_xlsx_file(excel_contents)
    try:
        subprocess.Popen(['C:\Program Files\Microsoft Office\Office14\excel.exe', excel_results_file])
    except FileNotFoundError:
        print('Excel not found at path: C:\Program Files\Microsoft Office\Office14\excel.exe')
        pass
    input('Press Enter to Close.')