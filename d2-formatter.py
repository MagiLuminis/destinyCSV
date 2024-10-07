import openpyxl
import json
import warnings
import datetime
import copy
import os
import sys


warnings.simplefilter(action='ignore', category=UserWarning)
def main(*args):
    #print(args[0][1])
    arguments = args[0]
    if len(arguments) < 2:
        input('No file name given\nEnter to quit...')
        quit()
    else:
        _file_name = arguments[1]

    script_path = os.path.dirname(os.path.abspath(__file__))
    working_directory = os.getcwd()
    os.chdir(script_path)
    working_directory = os.getcwd()
    
    _file_name = working_directory + '\\' +  _file_name
    workbook = openpyxl.load_workbook(filename=_file_name)
    sheet = workbook.active

    row_start = 9 #Starting Weapons Row (no headers)
    col_start = 4 #Starting Weapons Column
    list_of_weapons = []
    rolls = []
    weapon_template = {
        "name" : None,
        "sheet" : "perkbench",
        "pve" : {
            "masterwork" : [],
            "greatPerks" : [],
            "goodPerks" : []
            },
        "pvp" : {
            "masterwork" : [],
            "greatPerks" : [],
            "goodPerks" : []
        },
        "mnk" : False,
        "controller": False}
    for row in sheet.iter_rows(min_row=row_start,
                               min_col=col_start,
                               max_col=21): #Counting from Column A to #Masterwork Column
        
        info = copy.deepcopy(weapon_template)
        if row[7].value == 'Y':
            info['name'] = row[0].value.lower() #Weapon Name - Column - Starts from Column D
            masterwork = [row[17].value.lower()] #Masterwork - Column - Counting from #Weapon Name Column
            perks = []
            for r in row[12:17]: #First - Masterwork - Columns
                if r.value:
                    _text = r.value.lower()
                    if ',' in _text:
                        temp_ = _text.split(',')
                        temp_[1] = ''.join(temp_[1][1:])
                        perks.extend(temp_)
                    else:
                        perks.append(_text)
            info['pve']['masterwork'] = masterwork
            info['pve']['greatPerks'] = perks
            info['pvp']['masterwork'] = masterwork
            info['pvp']['greatPerks'] = perks
            info['mnk'] = True
            rolls.append(info)
            _info = copy.deepcopy(info)
            _info['mnk'] = False
            _info['controller'] = True
            rolls.append(_info)

    # finalize dates and dictionary for json
    date_json = datetime.datetime.now().strftime('%Y-%m-%d')
    date_file_name = datetime.datetime.now().strftime('%Y.%m.%d-%H.%M')
    whole_json = {
           "title" : "MagiLuminis",
           "date" : date_json,
           "manifestVersion" : "108125.22.08.29.2044-1-bnet.46276",
           "rolls" : rolls}

    #print(working_directory)
    #Add to open after D2GR- in case you need to have date/time -{date_file_name}
    with open(f'{working_directory}\\D2GR.json', 'w') as file:
        json_object = json.dumps(whole_json, indent='\t')
        file.write(json_object)
    #input("Enter")

    
if __name__ == '__main__':
    main(sys.argv)
    #input('Enter to Continue...') #Remove the hastag in case you need to see the result of the script
