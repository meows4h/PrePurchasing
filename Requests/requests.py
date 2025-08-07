# NOTE : 7/10/25
# is it possible to have the license section become a user # / copy amount section?
# adding cases for errors / warnings

# NOTE :
# if no license and/or permalink for ebook format, print errors to let user know that they need to fix that ; dont send out incomplete books
# ready to email status listings
# have list of changes to be done to the sheet if there are multiple errors i.e. missing permalink, malformed permalink, needing to add a new row if multiple IDs in it etc
# for scanned copies, whatever the number is for the CDL tag is the user number

import os
import pandas as pd
import sys
import configparser

# configure this script using the .ini file, should be much more readable this way
full_config = configparser.ConfigParser()
full_config.read('config.ini')
config = full_config['DEFAULT'] # reading the DEFAULT section
term = config['Term']
debug = config['Debug']
feedback = config['Feedback']
sheetname = config['SheetName']
skiplinks = config['SkipLinks']
curr_dir = os.path.dirname(__file__)

if debug == 'False': debug = False
if feedback == 'False': feedback = False
if skiplinks == 'False': skiplinks = False

if debug: print('Debug texts are active.')
if feedback: print('Sheet feedback is active.') 
else: print('Sheet feedback is not active. Please turn this on in the config to see suggested sheet changes.')

# checking for blank input directories
if config['InputDir'] is not None: 
    input_path = os.path.join(curr_dir, config['InputDir'], config['InputFile']) 
else:
    input_path = os.path.join(curr_dir, config['InputFile'])

if config['OutputDir'] is not None: 
    output_path = os.path.join(curr_dir, config['OutputDir'], config['OutputFile'])
else:
    output_path = os.path.join(curr_dir, config['OutputFile'])


# pulling original methods
def check_dir(directory:str, config):
    '''If the output path already exists, sends a check to the user to make sure
    they do not overwrite their already existing data.'''

    if os.path.exists(directory):
        user_inp = input(f'''File {config['OutputFile']} already exists.
Input "y" to continue & overwrite; "n" to exit.
Input: ''')
        
        if user_inp == 'y':
            pass
        else:
            sys.exit('\nExiting program.')


def duplicates_check():
    '''Asks the user if they would like to remove the duplicate ebook listings.'''

    user_inp = input(f'''Would you like to remove duplicate ebook listings for each professor?
Input "y" to remove duplicates; "n" to use data as is.
Input: ''')
    if user_inp == 'y':
        return True
    else:
        return False


def get_edition_str(row, row_name):
    '''Takes the current working row and checks it for the edition number, creating a string for it.'''

    edition_num = str(row[row_name]) if pd.isna(row[row_name]) == False else None
    th_list = ['4', '5', '6', '7', '8', '9', '0']

    if edition_num is not None:
        if edition_num[-1] == '1':
            edition_num += 'st'
        elif edition_num[-1] == '2':
            edition_num += 'nd'
        elif edition_num[-1] == '3':
            edition_num += 'rd'
        elif edition_num[-1] in th_list:
            edition_num += 'th'

        edition_num += ' Edition'
    
    return edition_num


def get_course_arr(row, row_name):
    '''Takes the current working row and extracts the course code and section, converting
    it into a single array for usage.'''

    course_code = ''
    last_char = False

    for char in row[row_name]:
        if char.isalpha():
            course_code += char
        elif char.isnumeric() and not last_char:
            course_code += ' '
            course_code += char
            last_char = True
        elif char.isnumeric():
            course_code += char
        elif char == '-':
            course_code += ' '

    course_arr = course_code.split(' ')

    # leaving out the section number, can revert this later if needed
    course_arr = [course_arr[0], course_arr[1]]

    return course_arr


def get_access_types(row):
    '''Takes the current working row and extracts all of the copy amounts and access links.'''

    # ebook, scan, physical book, streaming video, other
    format = row['Format']
    license = row['License (Electronic-Only)'] if not pd.isna(row['License (Electronic-Only)']) else ''
    permalink = row['PERMALINK'] if not pd.isna(row['PERMALINK']) else ''

    return_row = [format, license, permalink]

    return return_row


def get_access_email(book):
    '''This function takes in a single books access data and reads across it to create the email
    with the various links and listing for each title.'''

    scanned_appear = False
    phyiscal_appear = False
    listing = ''

    # ebooks
    if book[0] == 'ebook':
        if (book[1] == 'unlimited'):
            listing = f'<li><a href="{book[2]}">Ebook</a>: unlimited simultaneous users</li>'
        else:
            # adding some extra cases to prevent breaks when data is formatted weird
            if not book[1] == '':
                user_arr = book[1].split('-')
                user_str = 'users' if int(user_arr[0]) != 1 else 'user'
            else:
                user_arr = '???'
                user_str = 'users'

            listing = f'<li><a href="{book[2]}">Ebook</a>: {user_arr[0]} simultaneous {user_str}</li>'
        
    # scan / cdl
    elif book[0] == 'scan':
        scanned_appear = True
        if (book[1] == 'unlimited'):
            listing = f'<li><a href="{book[2]}">Scanned Book</a>: unlimited simultaneous users</li>'
        else:
            if not book[1] == '':
                user_arr = book[1].split('-')
                user_str = 'users' if int(user_arr[1]) != 1 else 'user'
            else:
                user_arr = '???'
                user_str = 'users'

            listing = f'<li><a href="{book[2]}">Scanned Book</a>: {user_arr[1]} simultaneous {user_str}</li>'

    # print books
    elif book[0] == 'physical book':
        phyiscal_appear = True
        # ????????
        print_copies = 1000000000 # no intake for this yet
        copy_str = 'copies' if print_copies != 1 else 'copy'
        listing = f'<li><a href="{book[2]}">Print</a>: [NUM COPIES] [{copy_str}] in Course Reserves</li>'

    # video
    elif book[0] == 'streaming video':
        listing = f'<li><a href="{book[2]}">Streaming Video</a></li>'

    # audiobooks
    elif book[0] == '????????':
        listing = f'<li><a href="{book[4][0]}">Audiobook</a> (limited users)</li>'

    # removing empty links from the html
    # if debug: print(f'LISTING DEBUG: {listing}')
    if '<a href="">' in listing:
        listing = listing.replace('<a href="">', '')
        listing = listing.replace('</a>', '')

    return listing, scanned_appear, phyiscal_appear


# read the data into the script
data = pd.read_excel(input_path, sheet_name=sheetname)

# check the output directory
check_dir(output_path, config)

# check if the user wants to remove duplicates
remove_duplicates = duplicates_check()

# set up the outline to grab all the necessary information
result_outline = {'Instructor' : [],
                  'Email' : [],
                  'Books' : [],
                  'Book Output' : [],
                  'Courses' : []
}

# format data output for output excel sheet
final_data = {'First Name' : [],
              'Last Name' : [],
              'Email' : [],
              'Book Output' : []
}


# looking across the whole input spreadsheet
for idx, row in data.iterrows():

    # checking if it is ready to email
    if row['Reserves Status'] != 'READY TO EMAIL':
        continue

    # checking format
    if pd.isna(row['Format']):
        if feedback: print(f'Error: Missing format for {row['Title'].title()} by {row['Author']} ({row['Instructor Name']})')
        continue

    # taking in the cell information
    instructor = str(row['Instructor Name'])
    email = str(row['Instructor Email'])

    # if the instructor or email doesnt exist, skip
    if instructor == '0' or email == '0' or instructor == '#N/A' or email == '#N/A' or pd.isna(row['Instructor Name']) or pd.isna(row['Instructor Email']):
        if feedback: print(f'Error: Missing instructor info for {row['Instructor Name']}, Book: {row['Title']}')
        continue

    # values to use for multiple instructors
    cont_flag = False # if there are multiple instructors and it cannot be parsed, this flag is used to skip this book listing
    instructor_arr = []
    email_arr = []

    # splits it into list between commas
    instructor = instructor.strip()
    name_arr = instructor.split(',')
    for name in name_arr:
        name = name.strip()

    ### NOTE: if anything needs an overhaul, it is the managing different multiple / single instructors input flow ###

    # this catches single names using a single space between them
    # i.e. "FirstName LastName", doesnt work with "FirstName LastName; FirstName..." etc, throws error otherwise
    if len(name_arr) == 1:
        name_arr = instructor.split(' ')
        if len(name_arr) == 2:
            first_name = name_arr[0]
            last_name = name_arr[1]
            instructor_arr.append(f'{first_name},{last_name}')
            email_arr.append(email)
        else:
            if feedback: print(f'Error: Unable to parse professor name into first / last (Possibly multiple names? Please split by commas) ({instructor}) E1') 
            continue
    
    # this catches single names using a single comma between them
    # i.e. "LastName, FirstName", however the check continues if there are more commas
    elif len(name_arr) == 2:
        first_name = name_arr[1]
        last_name = name_arr[0]
        instructor_arr.append(f'{first_name},{last_name}')
        email_arr.append(email)

    # this catches multiple names using the previous format, but with either a " and " or ; between them
    # i.e. "LastName, FirstName ; LastName, FirstName..." or "LastName, FirstName and LastName, FirstName and ..."
    # otherwise this throws an error if it cannot find it in this format
    # this also throws an error if emails are not explicitly listed as "email;email;email;..."
    elif len(name_arr) > 2:
        name_arr = instructor.split(';')
        if len(name_arr) == 1:
            name_arr = instructor.split(' and ')
        if len(name_arr) == 1:
            if feedback: print(f'Error: Unable to parse professor name into first / last (Possibly multiple names? Please split by commas) ({instructor}) E2')
            continue

        for idx, name in enumerate(name_arr):
            name = name.strip()
            temp_arr = name.split(',')
            if len(temp_arr) != 2:
                if feedback: print(f'Error: Unable to parse professor name into first / last (Possibly multiple names? Please split by commas) ({instructor}) E3')
                cont_flag = True
                break
            first_name = temp_arr[1].strip()
            last_name = temp_arr[0].strip()
            instructor_arr.append(f'{first_name},{last_name}')
            email_temp = email.split(';')
            if len(email_temp) <= 1:
                if feedback: print(f'Error: Unable to parse professor information, missing emails for multiple professors possibly? ({instructor})')
                cont_flag = True
                break
            email_arr.append(email_temp[idx])
        
        if cont_flag: continue

    ### NOTE: end of code that likely needs to be overhauled ###

    else:
        if feedback: print(f'Error: Unable to parse professor name into first / last (Possibly multiple names? Please split by commas) ({instructor}) E4') 
        continue

    # grabbing the edition number, course code string, and access types/links
    edition_num = get_edition_str(row, '(Edition)')
    course_arr = get_course_arr(row, 'Course')
    access_type = get_access_types(row)

    # compile baseline information + remove blank space
    title = row['Title'].title()
    title = title[:-1] if title[-1] == ' ' else title
    author = row['Author'].title()
    author = author[:-1] if author[-1] == ' ' else author
    book_info = [title, edition_num, author, access_type, course_arr, row['Date Pub.']]

    # iterating through all instructors in case of multiple
    for idx, instruct_temp in enumerate(instructor_arr):
    # add new instructor to the outline
        if instruct_temp not in result_outline['Instructor']:
            result_outline['Instructor'].append(instruct_temp)
            result_outline['Email'].append(email_arr[idx])
            result_outline['Books'].append([book_info])
            result_outline['Courses'].append([book_info[4]])

        # add any new information for existing instructors
        else:
            person_idx = result_outline['Instructor'].index(instruct_temp)
            result_outline['Books'][person_idx].append(book_info)

            if book_info[4] not in result_outline['Courses'][person_idx]:
                result_outline['Courses'][person_idx].append(book_info[4])


# taking the result outline and finalizing it into an output for email
# for each instructor found
for instructor in result_outline['Instructor']:

    # name is deconstructed according to the intake method earlier
    name_arr = instructor.strip().split(',')
    for name in name_arr:
        name = name.strip()
    #if debug: print(f'CHECKING NAME ARRAY: {name_arr}')
    first_name = name_arr[1]
    last_name = name_arr[0]

    idx = result_outline['Instructor'].index(instructor)

    email_str = ''
    scanned_appear = False
    phyiscal_appear = False

    # for each course the instructor is under
    for k, course in enumerate(result_outline['Courses'][idx]):

        course_str = '' # holding the string for the course book information
        
        # adding a break space if this is not the first subject found
        if k != 0:
            course_str += '<br>'

        course_str += f'<b>{course[0]} {course[1]}</b><br>Students can find all library materials by <a href="https://search.library.oregonstate.edu/discovery/search?query=any,contains,{course[0]}%20{course[1]}&tab=CourseReserves&search_scope=CourseReserves&vid=01ALLIANCE_OSU:OSU&lang=en&offset=0">searching course reserves for {course[0]} {course[1]}</a>.<br><ul>'
        book_list_str = '' # temp string to hold book info for the course, to be checked later

        # for each book in that course
        used_books = []
        for book in result_outline['Books'][idx]:

            # for readability
            book_title = book[0]
            book_edition = book[1]
            book_author = book[2]
            book_access = book[3]
            book_course = book[4]
            book_year = book[5]

            # skipping books not in the current course
            if book_course != course:
                continue
            
            # removing these repeated tags
            cleaner_list = [' (Cei)', '(Loose-Leaf)', 'Loose-Leaf','Ebook - Lifetime Duration', 'Ebook(5 Yr Access)', 'Ebook (Lifetime)', 'Ebook (180 days)', 'Ebook (150 days)', 'Ebook (120 days)', 'Ebook - Lifetime Access', 'Ebook -Lifetime Access', 'Ebook - Lifetime', 'Ebook - 180Days', 'Etext W/Connect Access Code', '[Qr]', '[Nbs]', '(Cei)', '(Ll)', 'W/1 Term Access Code Pkg', 'W/1 Year Access Code Pkg', 'W/2 Year Access Code Pkg', '1 Term Access Code', '1 Year Access Code', '2 Year Access Code', 'Ebook']
            if remove_duplicates:
                for phrase in cleaner_list:
                    book_title = book_title.replace(phrase.title(), '')

            # NOTE : Maybe fix authors with Mc- in the start of their name?
            book_author = book_author.replace('(Digital)', '')
            book_author = book_author.replace('(2)', '')

            # removing trailing whitespace
            while book_title[-1] == ' ':
                book_title = book_title[:-1]

            while book_author[-1] == ' ':
                book_author = book_author[:-1]

            # removing odd capitalization after apostrophes
            last_char = ''
            for i, character in enumerate(book_title):
                if last_char == "'":
                    temp_list = list(book_title)
                    temp_list[i] = book_title[i].lower()
                    book_title = ''.join(temp_list)
                last_char = character

            # removing odd 'O' capitalizations from the word 'of'
            if ' Of ' in book_title:
                book_title = book_title.replace(' Of ', ' of ')

            # skip if duplicate
            if book_title in used_books and remove_duplicates:
                continue

            # skip if missing permalink, unless skiplinks is true
            if pd.isna(book_access[2]):
                if feedback and skiplinks: 
                    print(f'Warning: Missing permalink for {book_title} by {book_author} ({last_name}) [Nothing listed.]')
                elif feedback: # make sure feedback is printed
                    print(f'Error: Missing permalink for {book_title} by {book_author} ({last_name}) [Nothing listed.]')
                    continue
                else: # still make sure to skip no matter what if skiplinks is false
                    continue

            elif 'https://search' != book_access[2][0:14]:
                if feedback and skiplinks: 
                    print(f'Warning: Missing permalink for {book_title} by {book_author} ({last_name}) [Not appropriate link.]')
                elif feedback:
                    print(f'Error: Missing permalink for {book_title} by {book_author} ({last_name}) [Not appropriate link.]')
                    continue
                else:
                    continue

            # adding in the book listing header
            book_list_str += f'<li><em>{book_title}</em>, {book_author}'
            book_list_str += f', {book_edition}' if book_edition is not None else ''
            book_list_str += f', {book_year}</li>' if not pd.isna(book_year) else '</li>'

            # add new email pieces
            access_email, scanned_appear, phyiscal_appear = get_access_email(book_access)
            if access_email != '':
                book_list_str += f'<ul>{access_email}</ul>'

            used_books.append(book_title)

        # if there are no eligible books for the course, skip course
        if book_list_str != '':
            course_str += book_list_str
            course_str += '</ul>'
            email_str += course_str
        else:
            course_str = ''
            continue

    # checking whether or not these specific cases have appeared for this professor
    if scanned_appear: 
        email_str += '<br>Scanned books are first come, first serve, for one hour at a time and use a waitlist. There is no limit to the number of renewals if no one is in the waitlist.'
    if scanned_appear and phyiscal_appear:
        email_str += '<br>'
    if phyiscal_appear:
        email_str += '<br>Physical copies are available for checkout at the Borrowing & Information desk for three hours at a time.'

    # if there is email content to send, add to the output list
    if email_str != '':
        final_data['Book Output'].append(email_str)
        final_data['First Name'].append(first_name)
        final_data['Last Name'].append(last_name)
        final_data['Email'].append(result_outline['Email'][idx])
    else:
        continue

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