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

# checking for blank input directories
if config['InputDir'] is not None: 
    input_path = os.path.join(curr_dir, config['InputDir'], config['InputFile']) 
else:
    input_path = os.path.join(curr_dir, config['InputFile'])

if config['OutputDir'] is not None: 
    output_path = os.path.join(curr_dir, config['OutputDir'], config['OutputFile'])
else:
    output_path = os.path.join(curr_dir, config['OutputFile'])

data = pd.read_excel(input_path)

# sends a check to the user to make sure data is not overwritten
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
                  'Books' : [],
                  'Book Output' : [],
                  'Courses' : []}

# looking across the whole input spreadsheet
for idx, row in data.iterrows():
    instructors = row['Instructor']
    instructor_arr = instructors.split('; ')

    # BIG NOTE: either ALL emails need to be listed properly! or possibly finding some hacked together way of checking the last name against the address but that isnt...
    emails = row['Instructor Email']
    emails_arr = emails.split('; ')

    # checking the access type
    access_type = [row['Access/Type'], row['Access/Type.1'], row['Access/Type.2']]
    if access_type[0] == 'not owned': # could add or statements to check other two... but this works for now
        continue

    # grabbing all the information for the book
    edition_num = str(int(row['Ed'])) if pd.isna(row['Ed']) == False else None

    # edition string
    if edition_num is not None:
        if edition_num[-1] == '1':
            edition_num += 'st'
        elif edition_num[-1] == '2':
            edition_num += 'nd'
        elif edition_num[-1] == '3':
            edition_num += 'rd'
        else:
            edition_num += 'th'

        edition_num += ' Edition'

    course_code = ''
    last_char = False

    for char in row['Course']:
        if char.isalpha():
            course_code += char
        elif char.isnumeric() and not last_char:
            course_code += ' '
            course_code += char
            last_char = True
        elif char.isnumeric():
            course_code += char

    book_info = [row['Title'].title(), edition_num, row['Author'].title(), [], [], course_code]

    # access link handler
    for k, access in enumerate(access_type):
        if pd.isna(access) == True:
            continue
        book_info[3].append(access) # adding the access type into an array
        book_info[4].append(row['Link']) if k == 0 else book_info[4].append(row[f'Link.{k}']) # adding the links in

    # professor data handler
    for j, person in enumerate(instructor_arr):

        if person not in result_outline['Instructor']:
            result_outline['Instructor'].append(person) # add new person
            result_outline['Email'].append(emails_arr[j]) # add their email to the array
            result_outline['Books'].append([book_info]) # add their own books array
            result_outline['Courses'].append([book_info[5]]) # add their own courses array

        else:
            person_idx = result_outline['Instructor'].index(person)
            result_outline['Books'][person_idx].append(book_info)

            if book_info[5] not in result_outline['Courses'][person_idx]:
                result_outline['Courses'][person_idx].append(book_info[5])

# format data output
final_data = {'First Name' : [],
              'Last Name' : [],
              'Email' : [],
              'Book Output' : []
}

for instructor in result_outline['Instructor']:
    name_arr = instructor.split(', ')
    first_name = name_arr[1]
    last_name = name_arr[0]
    idx = result_outline['Instructor'].index(instructor)

    final_data['First Name'].append(first_name)
    final_data['Last Name'].append(last_name)
    final_data['Email'].append(result_outline['Email'][idx])

    email_str = ''
    scanned_appear = False
    phyiscal_appear = False

    # the way this check is performed is really inefficient (for programming standards), so if program is slow, this is the first element to optimize as this is O(n^2) (if not worse)
    # this can be done in a way that is O(n) (reading the books list once) rather than reading it for every course there is
    # i forgot python says lists in dictionaries is bad (which is sound programming mind you) so i reworked the implementation based on that (this was originally going to use a hash, not lists)
    for k, course in enumerate(result_outline['Courses'][idx]):
        if k != 0:
            email_str += '<br>'
        
        course_arr = course.split(' ')

        email_str += f'<b>{course}</b><br>Students can find all library materials by <a href="https://search.library.oregonstate.edu/discovery/search?query=any,contains,{course_arr[0]}%20{course_arr[1]}&tab=CourseReserves&search_scope=CourseReserves&vid=01ALLIANCE_OSU:OSU&lang=en&offset=0">searching course reserves for {course}</a>.<br><ul>'
        for book in result_outline['Books'][idx]:
            if book[5] != course:
                continue

            email_str += f'<li><em>{book[0]}</em>, {book[2]}'
            email_str += f', {book[1]}' if book[1] is not None else ''
            email_str += '</li>' # NOTE: this is where publishing year goes, need spreadsheet update
            email_str += '<ul>'

            # holding all the strings so it is easier to iterate through them then just add them back later to the email_str
            ebook_list = ''
            scan_list = ''
            audio_list = ''
            video_list = ''
            print_list = ''
            dvd_list = ''

            # handling print in a special way since it can be either main collection or print reserves
            print_location = ''
            print_copies = 0

            for j, access_type in enumerate(book[3]):

                if access_type == 'unlimited':
                    ebook_list += f'<li><a href="{book[4][j]}">Ebook</a>: unlimited simultaneous users</li>'

                elif '-user' in access_type:
                    scanned_appear = True
                    user_arr = access_type.split('-')
                    user_str = 'users' if int(user_arr[0]) != 1 else 'user'
                    scan_list += f'<li><a href="{book[4][j]}">Scanned Book</a>: {user_arr[0]} simultaneous {user_str}</li>'

                elif access_type == 'CDL':
                    # NOTE: spreadsheet needs CDL user number update
                    scanned_appear = True
                    scan_list += f'<li><a href="{book[4][j]}">Scanned Book</a>: ??? simultaneous users</li>'

                elif access_type == 'audio':
                    audio_list += f'<li><a href="{book[4][j]}">Audiobook</a>: ???'

                # streaming video..? need example of what that type is

                elif access_type == 'print reserves':
                    print_copies += 1
                    if print_copies == 0:
                        print_location = book[4][j]
                
                elif access_type == 'main collection': 
                    print_copies += 1
                    if print_copies == 0:
                        print_location = book[4][j]

                # else:
                #     email_str += f'({access_type})'

                #email_str += '</li>'
            
            # NOTE: physical copies have no permalink accessible right now on the spreadsheet, needs update
            # NOTE: spreadsheet is also missing copy amounts so... yeah
            if print_copies > 0:
                phyiscal_appear = True
                print_list += f'<li><a href="">Print</a>: ??? copy/copies at {print_location}'

            access_listings = [ebook_list, scan_list, audio_list, video_list, print_list, dvd_list]
            for listing in access_listings:
                if listing != '':
                    email_str += listing
            
            email_str += '</ul>'


        email_str += '</ul>'

    if scanned_appear: 
        email_str += '<br>Scanned books are first come, first serve, for one hour at a time and use a waitlist. There is no limit to the number of renewals if no one is in the waitlist.'
    if phyiscal_appear:
        email_str += '<br>Physical copies are available for checkout at the Borrowing & Information desk for three hours at a time.'

    final_data['Book Output'].append(email_str)

# convert to dataframe object
result_data = pd.DataFrame(data=final_data)

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