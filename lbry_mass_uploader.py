import os
import pandas as pd
import time
from pprint import pprint
import logging
import requests


def logging_config():

    logfilename = 'log-' + 'lbry_mass_uploader' + '.txt'
    logging.basicConfig(filename=logfilename, level=logging.DEBUG,
                        format=' %(asctime)s-%(levelname)s-%(message)s')
    # set up logging to console
    console = logging.StreamHandler()
    console.setLevel(logging.INFO)
    # set a format which is simpler for console use
    formatter = logging.Formatter(' %(asctime)s-%(levelname)s-%(message)s')
    console.setFormatter(formatter)
    # add the handler to the root logger
    logging.getLogger('').addHandler(console)
    logger = logging.getLogger(__name__)
    

def clean_cmd():
    
    clear = lambda: os.system('cls')
    clear()


def get_next_file_datas():
    """
    :return: Dict.
    """

    df = pd.read_excel('list_upload.xlsx')
    mask_not_upload = df['upload_status'].isna()
    df_to_upload = df.loc[mask_not_upload, :].reset_index()
    qt_to_upload = df_to_upload.shape[0]
    if qt_to_upload == 0:
        return False
    else:
        logging.info(f"{qt_to_upload} files left")
        dict_next_file_datas_to_upload = df_to_upload.loc[0, :]
        return dict_next_file_datas_to_upload
        

def mark_file_as_request_upload(path_file):

    df = pd.read_excel('list_upload.xlsx')
    mask_file = df['path_file'].isin([path_file])
    df.loc[mask_file, 'upload_status'] = 1
    df.to_excel('list_upload.xlsx', index=False)
    
   
def save_txt(str_content, str_name):

    text_file = open(f"{str_name}.txt", "w")
    text_file.write(str_content)
    text_file.close()

    
def main():
    
    bid = '0.01'
    answer_qt = input('Upload how many files? ')
    int_answer_qt = int(answer_qt)
    for i in range(int_answer_qt):
        file_data = get_next_file_datas()
        print(file_data)
        a = requests.post("http://localhost:5279", 
                      json={"method": "stream_create", 
                            "params": {"bid": bid,
                                       "name": file_data['name_link_lbry'], 
                                       "title": file_data['title'],
                                       "description": file_data['description'],
                                       "file_path": file_data['path_file'],
                                       "thumbnail_url": file_data['thumbnail_url'],
                                       "tags": file_data['tags'].split(','),
                                       "validate_file": False, 
                                       "optimize_file": False, 
                                       "languages": file_data['languages'],
                                       "channel_account_id": [], 
                                       "funding_account_ids": [], 
                                       "preview": False, 
                                       "blocking": False}}).json()

        save_txt(str(a), 'return'+file_data['name_link_lbry'])
        pprint(a)
        mark_file_as_request_upload(file_data['path_file'])
    
   
if __name__ == "__main__":
    logging_config()
    main()