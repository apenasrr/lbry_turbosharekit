import pandas as pd
import os
import xlwings as xw    # to work with open excel files
import unidecode
import ast   
import pprint

def gen_report(path_dir):
    
    print(path_dir)
    l=[]
    for root, dirs, files in os.walk(path_dir):
        for file in files:        
            print(file)
            path_file = os.path.join(root, file)
            
            d={}
            d['path_file'] = path_file
            d['file_folder'] = root
            d['file_name'] = os.path.splitext(file)[0]
            d['file_size'] = os.path.getsize(path_file)
            d['title'] = ''
            d['name_link_lbry'] = ''
            d['description'] = ''
            d['thumbnail_url'] = ''
            d['tags'] = ''
            d['languages'] = "pt-BR"
            d['channel_id'] = ''
            d['upload_status'] = ''
            d['upload_datetime'] = ''
            l.append(d)
    df = pd.DataFrame(l)
    columns_order = ['path_file', 'file_folder', 'file_name', 'file_size', 
                     'title', 'name_link_lbry', 'description', 'thumbnail_url', 
                     'tags', 'languages', 'channel_id', 'upload_status', 
                     'upload_datetime']
    df = df.reindex(columns=columns_order)
    
    return df


def create_spreadsheet_upload():
    """
    Make a spreadsheet base to serve as input to massupload tool
    """
    
    path_dir = input('Paste the folder link where the files to upload: ')
    df = gen_report(path_dir)
    df.to_excel('list_upload.xlsx', index=False)
    print(f'\nComplement the metadata of the file uploads in report "list_upload.xlsx"')


def main():

    # ask if the user want to create spreadsheep or fill it
    list_options = ['Create spreadsheet_upload', 'Prefill the spreadsheet']
    question = 'Choose a process: '
    answer = choose_list_option(list_options, question)
    
    if answer == 'Create spreadsheet_upload':
        create_spreadsheet_upload()
        break_point = input('Type something to finish')
    elif answer == 'Prefill the spreadsheet':
        prefill_file()
        break_point = input('Type something to finish')
    else:
        print('\nAnswer invalid.\n')
        break_point = input('Type something to restart')
        clean_cmd()
        main()
        return
    

def fill_seq_column(sheet_obj, standard_str, qt_lines, col_n):

    for index_ini in range(1, qt_lines+1):
        sheet_obj.cells(index_ini+1, col_n).value = \
            standard_str + '_%02d-%02d' % (index_ini, qt_lines)


def normalize_string_to_link(string_actual):

    string_new = unidecode.unidecode(string_actual)
    
    for c in r"!@#$%^&*()[]{};:,./<>?\|`~-=_+":
        string_new = string_new.translate({ord(c): "_"})
        
    string_new = string_new.replace(' ', '_')
    string_new = string_new.replace('___', '_')
    string_new = string_new.replace('__', '_')
    
    return string_new
    

def autofill_title(sheet_obj, qt_lines, col_title):

    standard_str = input('Inform a standard Title: ')
    fill_seq_column(sheet_obj, standard_str, qt_lines, col_title)


def autofill_lbrylink(sheet_obj, qt_lines, col_title, col_lbrylink):

    print('\n1-From titles')
    print('2-From a standard sequenced name\n')
    answer_link_lbry_method = input('Which link building method ' +
                                    'do you prefer? ')
                                    
    if answer_link_lbry_method == '1':
        
        for line_index in range(2, qt_lines+2):
            string_title = sheet_obj.cells(line_index, col_title).value
            string_link_formated = normalize_string_to_link(string_title)
            
            sheet_obj.cells(line_index,
                            col_lbrylink).value = string_link_formated
            

    if answer_link_lbry_method == '2':
    
        standard_str = input('Inform a standard Lbry Link: ')
        string_link_formated = normalize_string_to_link(standard_str)
        fill_seq_column(sheet_obj, string_link_formated, 
                        qt_lines, col_lbrylink)
    
  
def get_txt_content(file_path):

    file_obj = open(file_path, 'r', encoding='utf-8')
    list_file_content = file_obj.readlines()
    file_obj.close()
    
    str_file_content = ''.join(list_file_content)
    
    return str_file_content
    

def autofill_description(sheet_obj, qt_lines, col_description):

    input('Please make sure the file "description.txt" ' + 
          'is properly filled.\nPress to continue')
    
    file_path = 'description.txt'
    str_description = get_txt_content(file_path)
    for index_ini in range(1, qt_lines+1):
        sheet_obj.cells(index_ini+1, col_description).value = str_description


def autofill_thumbnail_url(sheet_obj, qt_lines, col_thumbnail_url):
    
    str_thumbnail_url = input('Inform a standard thumbnail url: ')
    for index_ini in range(1, qt_lines+1):
        sheet_obj.cells(index_ini+1, col_thumbnail_url).value = str_thumbnail_url
        

def autofill_tags(sheet_obj, qt_lines, col_tags):

    print('Inform a list of tags (max 5), separated by comma (,). ' +
          '\ne.g.: music,beatles,live')
    str_tags = input('Tags: ')
    tags_not_allow = [';', '#', '/']
    for tag_not_allow in tags_not_allow:
        if tag_not_allow in str_tags:
            print('\nTags must be separated by comma (,).' + 
                  '\ne.g.: music,beatles,live\n')
            autofill_tags(sheet_obj, qt_lines, col_tags)
            return
            
    str_tags_formated = str_tags.replace(', ', ',').replace(' ,', ',')
    for index_ini in range(1, qt_lines+1):
        sheet_obj.cells(index_ini+1, col_tags).value = str_tags_formated


def ask_choose_lang():
    
    d_lang = {}
    d_lang["Arabic"] = "ar"
    d_lang["Chinese"] = "zh"
    d_lang["Croatian"] = "hr"
    d_lang["Czech"] = "cs"
    d_lang["Dutch"] = "nl"
    d_lang["English"] = "en"
    d_lang["Finnish"] = "fi"
    d_lang["French"] = "fr"
    d_lang["German"] = "de"
    d_lang["Greek"] = "el"
    d_lang["Hindi"] = "hi"
    d_lang["Indonesian"] = "id"
    d_lang["Italian"] = "it"
    d_lang["Japanese"] = "jp"
    d_lang["Kannada"] = "kn"
    d_lang["Khmer"] = "km"
    d_lang["Korean"] = "ko"
    d_lang["Malay"] = "ms"
    d_lang["Norwegian"] = "no"
    d_lang["Polish"] = "pl"
    d_lang["Portuguese"] = "pt"
    d_lang["Romanian"] = "ro"
    d_lang["Russian"] = "ru"
    d_lang["Spanish"] = "es"
    d_lang["Thai"] = "th"
    d_lang["Turkish"] = "tr"
    d_lang["Vietnamese"] = "vi"
    
    # ask to choose language
    list_lang = []
    for key in d_lang:
        list_lang.append(key)

    for index, lang in enumerate(list_lang, 1):
        print(f'{index}-{lang}')
        
    answer_lang_num = int(input('Choose a standard language: '))
    answer_lang = list_lang[answer_lang_num-1]
    cod_lang = d_lang[answer_lang]
    
    return cod_lang
    
    
def autofill_languages(sheet_obj, qt_lines, col_languages):
    
    # check on config lang and ask if the user want use it
    path_file = os.path.join('config', 'config.txt')
    variable_name='lang'
    dict_return = handle_config_file(path_file, variable_name=variable_name, 
                                      set_value=None, parse=True)
    lang_default = dict_return[variable_name][0]
    set_default = input(f'(None for yes) Do you want '+
                        f'select "{lang_default}" language? ')
    if set_default == '':
        cod_lang = lang_default
    else:
        # ask for choose new lang
        cod_lang = ask_choose_lang()    
        handle_config_file(path_file, variable_name=variable_name, 
                           set_value=cod_lang)
    
    for index_ini in range(1, qt_lines+1):
        sheet_obj.cells(index_ini+1, col_languages).value = cod_lang
    

def choose_list_option(list_options, question):

    print('')
    for index, option in enumerate(list_options, 1):
        print(f'{index}-{option}')
    answer_str = input('\n' + question)
    if answer_str == '':
        return
    index_answer_int = int(answer_str) - 1
    value_choose = list_options[index_answer_int]
    
    return value_choose


def prefill_file():
    
    answer_autofill = input('(Default No) Would you like to do autofill? ')
    if answer_autofill == '':
        return
        
    full_file = r'list_upload.xlsx'

    wb = xw.Book(full_file)
    sheet_obj = wb.sheets[0]
    last_cell = sheet_obj.range('A1').current_region.last_cell.row
    qt_lines = last_cell - 1
    col_title = 5
    col_lbrylink = 6
    col_description = 7
    col_thumbnail_url = 8
    col_tags = 9
    col_languages = 10
    col_channel_id = 11
    
    keep_fill = True
    while keep_fill:
        list_cols = ['title', 'name_link_lbry', 'description', 'thumbnail_url', 
                     'tags', 'languages', 'channel_id']
        question = 'Which column would you like to autofill? '
        answer_option = choose_list_option(list_cols, question)
        
        if answer_option == 'title':    
            autofill_title(sheet_obj, qt_lines, col_title)
                    
        if answer_option == 'name_link_lbry':    
            autofill_lbrylink(sheet_obj, qt_lines, col_title, col_lbrylink)
        
        if answer_option == 'description':    
            autofill_description(sheet_obj, qt_lines, col_description)

        if answer_option == 'thumbnail_url':    
            autofill_thumbnail_url(sheet_obj, qt_lines, col_thumbnail_url)

        if answer_option == 'tags':    
            autofill_tags(sheet_obj, qt_lines, col_tags)

        if answer_option == 'languages':    
            autofill_languages(sheet_obj, qt_lines, col_languages)
            
        if answer_option == 'channel_id':    
            autofill_channel_id(sheet_obj, qt_lines, col_channel_id)
        
        answer_continue = input('\n(None for yes) Continue to prefill?\n')
        if answer_continue != '':
            keep_fill = False
        


def autofill_channel_id(sheet_obj, qt_lines, col_channel_id):
    
    path_file = os.path.join('config', 'config.txt')
    variable_name='channel'
    dict_value = handle_config_file(path_file, variable_name=variable_name, 
                                    parse=True)
    
    dict_channel = dict_value[variable_name]
    print('')
    list_key = []
    for index, key in enumerate(dict_channel, 1):
        print(f'{index}-{key}')
        list_key.append(key)
        
    answer_channel_num = input('(None to add a new) Choose a channel: ')
    #TODO possibility to exclude a channel
    
    # possibility to register another channel
    if answer_channel_num == '':
        new_channel_name = input('Inform the channel name: ')
        
        new_channel_id = input('Inform the channel id: ')
        
        d_new_channel = {}
        d_new_channel[new_channel_name] = new_channel_id
        
        # Use Handler include new channel in config file
        handle_config_file(path_file, variable_name=variable_name, 
                           set_value=d_new_channel)
        
    else:
        key_choose = list_key[int(answer_channel_num)-1]
        channel_id_choose = dict_channel[key_choose]
        # Fill channel id
        for index_ini in range(1, qt_lines+1):
            sheet_obj.cells(index_ini+1, col_channel_id).value = channel_id_choose
      
    
def config_file_parser_values(list_found, variable_name):
    
    list_found_parsed = []
    dict_build = {}
    dict_values = {}
    for item in list_found:
        item_parsed = ast.literal_eval(item)
        if isinstance(item_parsed, dict):
            dict_values.update(item_parsed)
        else:
            list_found_parsed.append(item_parsed)

    if len(dict_values)!=0:
        dict_build[variable_name] = dict_values
        if len(list_found_parsed) != 0:
            dict_build[variable_name]['others'] = list_found_parsed
    else:
        dict_build[variable_name] = list_found_parsed
    return dict_build


def handle_config_file(path_file, variable_name, set_value=None, 
                       parse=False):

    def get_updated_line(variable_name, set_value):
        if isinstance(set_value, dict):
            set_value_parsed = set_value
        else:
            set_value_parsed = set_value
            
        if isinstance(set_value_parsed, dict):
            updated_line = f"{variable_name}={set_value_parsed}\n"
        else:
            updated_line = f"{variable_name}='{set_value_parsed}'\n"
        return updated_line
    
    def get_str_value(line):
        line_components = line.split('=')
        str_value = line_components[1]
        str_value = str_value.replace("\n", '')
        return str_value

    def get_item_parsed(line):
        str_value = get_str_value(line)
        item_parsed = ast.literal_eval(str_value)
        return item_parsed
        
    def value_is_dict(item_parsed):
        is_dict = isinstance(item_parsed, dict)
        return is_dict

    def is_same_key(item_parsed, set_value):
        key_item_parsed = next(iter(item_parsed))
        key_set_value = next(iter(set_value))
        same_key = key_item_parsed == key_set_value
        return same_key
        
    config_file = open(path_file, 'r+')
    content_lines = []

    list_found = []
    dont_found = True
    if set_value:
        updated_line = get_updated_line(variable_name, set_value)
        for line in config_file:
            if f'{variable_name}=' in line:
                item_parsed = get_item_parsed(line)
                if value_is_dict(item_parsed):
                    if is_same_key(item_parsed, set_value):
                        dont_found = False
                        content_lines.append(updated_line)
                    else:
                        content_lines.append(line)
                else:
                    dont_found = False
                    content_lines.append(updated_line)
            else:
                content_lines.append(line)
              
        if dont_found:
            
            # include variable_name and value at botton of file
            content_lines.append(updated_line)

        # save and finish file
        config_file.seek(0)
        config_file.truncate()
        config_file.writelines(content_lines)
        config_file.close()
        
    else:
        for line in config_file:
            if f'{variable_name}=' in line:
                str_value = get_str_value(line)
                list_found.append(str_value)
        # finish file and return value
        config_file.close()
        if parse:
            dict_build = config_file_parser_values(list_found, variable_name)
            return dict_build
        else:
            dict_build = {}
            dict_build[variable_name] = list_found 
            return dict_build
                
        
def clean_cmd():
    
    clear = lambda: os.system('cls')
    clear()


if __name__ == "__main__":
    main()