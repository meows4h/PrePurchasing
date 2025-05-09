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
    edition_num = str(int(row['Ed'])) if pd.isna(row['Ed']) == False else edition_num = None

    if edition_num[-1] == '1':
        edition_num += 'st'
    elif edition_num[-1] == '2':
        edition_num += 'nd'
    elif edition_num[-1] == '3':
        edition_num += 'rd'
    elif edition_num is not None:
        edition_num += 'th'

    book_string = f'{row['Title']}'
    book_string += f'{edition_num} Edition' if edition_num is not None else ''
    book_string += f'by {row['Author']} for {row['Course']}, available from '

    book_info = [row['Title'].title(), edition_num, row['Author'].title(), [], [], [row['Course']]]

    for k, access in enumerate(access_type):
        if pd.isna(access) == True:
            continue
        book_info[3].append(access) # adding the access type into an array
        book_info[4].append(row['Link']) if k == 0 else book_info[4].append(row[f'Link.{k}']) # adding the links in

        if k > 0:
            book_string += ' It is also available from '

        if access == 'unlimited':
            link_str = 'Link' if k == 0 else link_str = f'Link.{k}'
            book_string += f'{row[link_str]} with an {access} user license.'

        elif access == 'CDL':
            # NOTE: This is supposed to have a specific number of users -- current version of the spreadsheet just lists CDL
            link_str = 'Link' if k == 0 else link_str = f'Link.{k}'
            book_string += f'{row[link_str]}. Online texts are first come, first serve, for one hour at a time, and are renewable so long as no one is on the waitlist.'

        # TODO: other access types
    
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