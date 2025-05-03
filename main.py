import os
import pandas as pd
import sys
import configparser

# configure this script using the .ini file, should be much more readable this way
full_config = configparser.ConfigParser()
full_config.read('config.ini')
config = full_config['DEFAULT'] # reading the DEFAULT section -- might change later to have multiple configs
term = config['Term']
curr_dir = os.path.dirname(__file__)

# checking for blank input
if config['InputDir'] is not None: 
    input_path = os.path.join(curr_dir, config['InputDir'], config['InputFile']) 
else:
    input_path = os.path.join(curr_dir, config['InputFile'])

if config['OutputDir'] is not None: 
    output_path = os.path.join(curr_dir, config['OutputDir'], config['OutputFile'])
else:
    output_path = os.path.join(curr_dir, config['OutputFile'])

# leftover debug print, for finding the current paths for inputs and outputs, double checking against config
# print(f'IP: {input_path}; OP: {output_path}; I: {config['InputDir']}; O: {config['OutputDir']}; IF: {config['InputFile']}; OF: {config['OutputFile']}')

def process_sheet(path):
    '''description for later'''
    sheet = pd.read_excel(path) #other cleaning up can be performed here
    return sheet

data = process_sheet(input_path)

def check_dir(directory:str):
    '''description for later'''
    if os.path.exists(output_path):
        user_inp = input(f'''File {config['OutputFile']} already exists.
                        Input "y" to continue & overwrite; "n" to exit.
                        Input: ''')
        
        if user_inp == 'y':
            pass
        else:
            sys.exit('\nExiting program.')

check_dir(output_path)

result_outline = {'Instructor' : [],
                  'Email' : [],
                  'Books' : []}

for idx, row in data.iterrows():
    instructors = row['Instructor']
    instructor_arr = instructors.split('; ')

    # BIG NOTE: either ALL emails need to be listed properly !!! or possibly finding some hacked together way of checking the last name against the address but that isnt...
    emails = row['Instructor Email']
    emails_arr = emails.split('; ')

    # checking the access type
    access_type = [row['Access/Type'], row['Access/Type.1'], row['Access/Type.2']]
    if access_type[0] == 'not owned': # could add or statements to check other two... but this works for now
        continue

    # grabbing all the information for the book
    book_info = [row['Title'].title()]
    book_info.append(row['Ed']) if f'{row['Ed']}' != 'nan' else book_info.append('N/A')
    book_info = [row['Title'].title(), row['Ed'], row['Author'].title(), [], []]
    for k, access in enumerate(access_type):
        if f'{access}' == 'nan':
            continue
        book_info[3].append(access) # adding the access type into an array
        book_info[4].append(row['Link']) if k == 0 else book_info[4].append(row[f'Link.{k}']) # adding the links in
    
    # checking if the instructor already exists and handling data accordingly
    for j, person in enumerate(instructor_arr):
        if person not in result_outline['Instructor']:
            result_outline['Instructor'].append(person)
            result_outline['Email'].append(emails_arr[j])
            result_outline['Books'].append([book_info])
        else:
            person_idx = result_outline['Instructor'].index(person)
            result_outline['Books'][person_idx].append(book_info)

# convert to dataframe object
result_data = pd.DataFrame(data=result_outline)

# open an xlsxwriter to be able to format the frame into a table
writer = pd.ExcelWriter(output_path, engine='xlsxwriter')
result_data.to_excel(writer, sheet_name='email list', startrow=1, header=False, index=False)

workbook = writer.book
worksheet = writer.sheets['email list']

(max_row, max_col) = result_data.shape

column_settings = []
for header in result_data.columns:
    column_settings.append({'header': header})

worksheet.add_table(0, 0, max_row,max_col - 1, {'columns': column_settings})

worksheet.set_column(0, max_col - 1, 12)
writer.close() # close and output

# result_data.to_excel(output_path) # writing the file to the output

# notes for next time
# we are looking to split this up into a sheet that organizes data by instructor
# so we can feed this to an automation program

# each instructor will have a name, associated email, and possibly notified tab
# each instructor will have a book
# each book will have a course (and section + enrollment), link, and availability (and access type)
# possibly look at including date of which things were modified or changed

# we need to see how this data gets sorted in pandas first, then go from there
# maybe the sheet can include the finished email instead? who knows