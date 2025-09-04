# NOTE : how to use formatting wise
# format examples
# Reserve Status : just looking for the single input
# i.e. "READY TO EMAIL"

# Course : spaces do not affect input, slashes indicate multiple classes or numbers
# i.e. "MTH 251", "MTH251", "MTH251Z", "MTH 251/253", "BI445/BI 545"

# Title : any string
# Author : any string
# Date Pub. : year ; i.e. 2005, 1994, etc

# Edition : a number is easiest, but only include one, any pieces like "nd" get removed
# i.e. 2, 6, 12, 3rd, 5th, 17th

# ISBN : this value is not read by this program as of this comment
# Format : this list -> ebook, physical book, scan, dvd, streaming video, audiobook, other
# Total Print Copies : just a number

# License (Electronic-Only) : just pick from the list
# i.e. single-user, 2-user, 3-user, etc.
# CDL-1, CDL-2, CDL-3, etc.
# OverDrive (OC/OU), OverDrive (Other)

# PERMALINK : needs a proper search.library link

# The remaining values are not processed as of this time


import os
import pandas as pd
import sys
import configparser

# configure this script using the .ini file
full_config = configparser.ConfigParser()
full_config.read('config.ini')
config = full_config['DEFAULT']  # reading the DEFAULT section
term = config['Term']
debug = config['Debug']
feedback = config['Feedback']
sheetname = config['SheetName']
skiplinks = config['SkipLinks']
remove_duplicates = config['RemoveDuplicateBooks']
curr_dir = os.path.dirname(__file__)

if debug == 'False': debug = False
if feedback == 'False': feedback = False
if skiplinks == 'False': skiplinks = False
if remove_duplicates == 'False': remove_duplicates = False

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


def get_edition_str(row, row_name):
    '''Takes the current working row and checks it for the edition number,
       creating a string for it.'''

    edition_num = str(row[row_name]) if pd.isna(row[row_name]) is False else None
    th_list = ['4', '5', '6', '7', '8', '9', '0']
    replace_list = ['st', 'nd', 'rd', 'th', 'ed', 'Edition']

    # double checking that it is not an invalid string to parse
    double_check = False

    if edition_num is not None:

        # removing weird inputs for editions
        for replacement in replace_list:
            edition_num = edition_num.replace(replacement, '')

        if ' ' in edition_num:
            double_check = True

        if edition_num[-1] == '1':
            edition_num += 'st'
        elif edition_num[-1] == '2':
            edition_num += 'nd'
        elif edition_num[-1] == '3':
            edition_num += 'rd'
        elif edition_num[-1] in th_list:
            edition_num += 'th'

        edition_num += ' Edition'

    if double_check:
        edition_num = None
        if feedback: print(f'WARNING: Book has odd edition number, {row["Title"]} by {row["Author"]} (skipping adding edition number)')
    elif edition_num is None:
        if feedback: print(f'WARNING: Book missing edition number, {row["Title"]} by {row["Author"]} (skipping adding edition number)')

    return edition_num


def get_course_arr(row, row_name):
    '''Takes the current working row and extracts the course code(s),
       converting it into a single array for usage.'''

    # if empty, return empty
    if pd.isna(row[row_name]):
        return ''

    course_code = ''
    course_arr = [[], []]  # list 1 stores subject, list 2 stores number
    subject_done = False

    for idx, char in enumerate(row[row_name]):

        # make sure each letter gets added
        if char.isalpha():
            course_code += char

        # if the subject piece is just finishing, add a space
        elif char.isnumeric() and not subject_done:
            course_code += ' '
            course_code += char
            subject_done = True

        # add each number
        elif char.isnumeric():
            course_code += char

        # if a slash is present, add the individual course
        elif char == '/' or char == '\\':
            course_temp = course_code.split(' ')
            course_arr[0].append(course_temp[0])
            course_arr[1].append(course_temp[1])
            subject_done = False

            # save the subject if it is two number codes
            # otherwise remove it
            if len(row[row_name]) > idx + 1:
                if row[row_name][idx + 1].isnumeric():
                    course_code = course_temp[0]
                else:
                    course_code = ''

    course_temp = course_code.split(' ')
    course_arr[0].append(course_temp[0])
    course_arr[1].append(course_temp[1])

    return course_arr


def get_access_types(row):
    '''Takes the current working row and extracts
       all of the copy amounts and access links.'''

    # ebook, scan, physical book, streaming video, other
    format = row['Format']
    license = row['License (Electronic-Only)'] if not pd.isna(row['License (Electronic-Only)']) else ''
    permalink = row['PERMALINK'] if not pd.isna(row['PERMALINK']) else ''
    print_copies = row['Total Print Copies'] if not pd.isna(row['Total Print Copies']) else ''

    return_row = [format, license, permalink, print_copies]

    return return_row


def get_access_email(book):
    '''This function takes in a single books access data and reads across it
       to create an email with the various links and listing for each title.'''

    scanned_appear = False
    phyiscal_appear = False
    listing = ''

    # ebooks
    if book[0] == 'ebook':
        if (book[1] == 'unlimited'):
            listing = f'<li><a href="{book[2]}">Ebook</a>: unlimited simultaneous users</li>'
        elif (book[1] == 'OverDrive (OC/OU)' or book[1] == 'OverDrive (Other)'):
            listing = f'<li><a href="{book[2]}">Ebook</a>: OverDrive license</li>'
        else:
            # extra cases to prevent breaks when data is formatted weird
            if book[1] == 'single-user':
                user_arr = '1'
                user_str = 'user'
            elif not book[1] == '':
                user_arr = book[1].split('-')
                user_str = 'users' if int(user_arr[0]) != 1 else 'user'
            else:
                user_arr = '???'
                user_str = 'users'

            listing = f'<li><a href="{book[2]}">Ebook</a>: {user_arr[0]} simultaneous {user_str}</li>'

    # ebooks cdl replace?
    # it is listed as an option but not implementing anything for it

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
        print_copies = int(book[3])
        copy_str = 'copies' if print_copies != 1 else 'copy'
        listing = f'<li><a href="{book[2]}">Print</a>: {print_copies} {copy_str} in Course Reserves</li>'

    # video
    elif book[0] == 'streaming video':
        listing = f'<li><a href="{book[2]}">Streaming Video</a></li>'

    # dvd
    elif book[0] == 'dvd':
        listing = f'<li><a href="{book[2]}">DVD</a></li>'

    # audiobooks
    elif book[0] == 'audiobook':
        listing = f'<li><a href="{book[2]}">Audiobook</a> (limited users)</li>'

    # other
    elif book[0] == 'other':
        listing = f'<li><a href="{book[2]}">Alternative Access Type</a></li>'

    # removing empty links from the html
    if '<a href="">' in listing:
        listing = listing.replace('<a href="">', '')
        listing = listing.replace('</a>', '')

    return listing, scanned_appear, phyiscal_appear


def clean_book_info(title, author):

    # removing these repeated tags
    cleaner_list = [' (Cei)', '(Loose-Leaf)', 'Loose-Leaf', 'Ebook - Lifetime Duration', 'Ebook(5 Yr Access)', 'Ebook (Lifetime)', 'Ebook (180 days)', 'Ebook (150 days)', 'Ebook (120 days)', 'Ebook - Lifetime Access', 'Ebook -Lifetime Access', 'Ebook - Lifetime', 'Ebook - 180Days', 'Etext W/Connect Access Code', '[Qr]', '[Nbs]', '(Cei)', '(Ll)', 'W/1 Term Access Code Pkg', 'W/1 Year Access Code Pkg', 'W/2 Year Access Code Pkg', '1 Term Access Code', '1 Year Access Code', '2 Year Access Code', 'Ebook']
    for phrase in cleaner_list:
        title = title.replace(phrase.title(), '')

    # NOTE : Maybe fix authors with Mc- in the start of their name?
    # i.e. currently McDonald looks like Mcdonald
    author = author.replace('(Digital)', '')
    author = author.replace('(2)', '')

    # removing odd capitalization after apostrophes
    last_char = ''
    for i, character in enumerate(title):
        if last_char == "'":
            temp_list = list(title)
            temp_list[i] = title[i].lower()
            title = ''.join(temp_list)
        last_char = character

    # removing/fixing other odd capitalizations
    if ' Of ' in title:
        title = title.replace(' Of ', ' of ')

    if ' A ' in title:
        title = title.replace(' A ', ' a ')

    if ' Ii ' in title:
        title = title.replace(' Ii ', ' II ')

    if ' Iii ' in title:
        title = title.replace(' Iii ', ' III ')

    if ' Ii' in title:
        title = title.replace(' Ii', ' II')

    if ' Iii' in title:
        title = title.replace(' Iii', ' III')

    title = title.strip()
    author = author.strip()

    return title, author


# read the data into the script
data = pd.read_excel(input_path, sheet_name=sheetname)

# check the output directory
check_dir(output_path, config)

# set up the outline to grab all the necessary information
result_outline = {'Email': [],
                  'Books': [],
                  'Book Output': [],
                  'Courses': []}

# format data output for output excel sheet
final_data = {'Email': [],
              'Book Output': []}


# looking across the whole input spreadsheet
for idx, row in data.iterrows():

    # checking if it is ready to email
    if row['Reserves Status'] != 'READY TO EMAIL':
        continue

    # checking format
    if pd.isna(row['Format']):
        if feedback: print(f'ERROR: Missing format for {row['Title'].title()} by {row['Author']} ({row['Instructor Name']})')
        continue

    # taking in the cell information
    instructor = str(row['Instructor Name'])
    email = str(row['Instructor Email'])

    # if the instructor or email doesnt exist, skip
    if email == '0' or email == '#N/A' or pd.isna(row['Instructor Email']):
        if feedback: print(f'ERROR: Missing email info for {instructor}, Book: {row['Title']}')
        continue

    # grabbing email(s), checks for , and ; seperation style
    email_arr = email.split(',')
    if len(email_arr) == 1:
        email_arr = email.split(';')

    for address in email_arr:
        address = address.strip()

    # grabbing the edition number, course code string, and access types/links
    edition_num = get_edition_str(row, '(Edition)')
    course_arr = get_course_arr(row, 'Course')
    access_type = get_access_types(row)

    # title
    if not pd.isna(row['Title']):
        title = row['Title'].title()
        title = title.strip()
    else:
        if feedback: print(f'ERROR: Missing book name for course(s): {course_arr}')
        continue

    # author
    if not pd.isna(row['Author']):
        author = row['Author'].title()
        author = author.strip()
    else:
        if feedback: print(f'ERROR: Missing author for {title}')
        continue

    # publishing year
    year = str(row['Date Pub.']).strip()
    if len(year) != 4:
        year = None
        if feedback: print(f'WARNING: {title} by {author} missing a clear publishing year (should just be a four digit year)')

    # catching the course error
    if course_arr == '':
        if feedback: print(f'ERROR: Missing associated course code for {title} by {author}')
        continue

    # iterating through all emails in case of multiple
    for idx, email_temp in enumerate(email_arr):

        # iterating through each course in case of multiple courses
        for kdx in range(len(course_arr[0])):
            course_code = [course_arr[0][kdx], course_arr[1][kdx]]
            book_info = [title, edition_num, author, access_type, course_code, year]

            # add new email to the outline
            if email_temp not in result_outline['Email']:
                result_outline['Email'].append(email_arr[idx])
                result_outline['Books'].append([book_info])
                result_outline['Courses'].append([course_code])

        # add any new information for existing emails
            else:
                person_idx = result_outline['Email'].index(email_temp)
                result_outline['Books'][person_idx].append(book_info)

                if course_code not in result_outline['Courses'][person_idx]:
                    result_outline['Courses'][person_idx].append(course_code)


# taking the result outline and finalizing it into an output for email
# for each email found
for email in result_outline['Email']:

    idx = result_outline['Email'].index(email)

    email_str = ''
    scanned_appear = False
    phyiscal_appear = False

    # for each course the email is under
    for k, course in enumerate(result_outline['Courses'][idx]):

        course_str = ''  # holding the string for the course book information

        # adding a break space if this is not the first subject found
        if k != 0:
            course_str += '<br>'

        course_str += f'<b>{course[0]} {course[1]}</b><br>Students can find all library materials by <a href="https://search.library.oregonstate.edu/discovery/search?query=any,contains,{course[0]}%20{course[1]}&tab=CourseReserves&search_scope=CourseReserves&vid=01ALLIANCE_OSU:OSU&lang=en&offset=0">searching course reserves for {course[0]} {course[1]}</a>.<br><ul>'
        book_list_str = ''  # temp string to hold book info for the course, to be checked later

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

            book_title, book_author = clean_book_info(book_title, book_author)

            # skip if duplicate
            if book_title in used_books and remove_duplicates:
                continue

            # skip if missing permalink, unless skiplinks is true
            if pd.isna(book_access[2]):
                if feedback and skiplinks: 
                    print(f'WARNING: Missing permalink for {book_title} by {book_author} ({email}) [Nothing listed.]')
                elif feedback:  # make sure feedback is printed
                    print(f'ERROR: Missing permalink for {book_title} by {book_author} ({email}) [Nothing listed.]')
                    continue
                else:  # still make sure to skip no matter what if skiplinks is false
                    continue

            elif 'https://search' != book_access[2][0:14]:
                if feedback and skiplinks:
                    print(f'WARNING: Missing permalink for {book_title} by {book_author} ({email}) [Not appropriate link.]')
                elif feedback:
                    print(f'ERROR: Missing permalink for {book_title} by {book_author} ({email}) [Not appropriate link.]')
                    continue
                else:
                    continue

            if book_access[0] == 'physical book' and book_access[3] == '':
                if feedback:
                    print(f'ERROR: Missing physical book count for {book_title} by {book_author} ({email})')
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

    # checking whether or not these specific cases have appeared for this email
    if scanned_appear:
        email_str += '<br>Scanned books are first come, first serve, for one hour at a time and use a waitlist. There is no limit to the number of renewals if no one is in the waitlist.'
    if scanned_appear and phyiscal_appear:
        email_str += '<br>'
    if phyiscal_appear:
        email_str += '<br>Physical copies are available for checkout at the Borrowing & Information desk for three hours at a time.'

    # if there is email content to send, add to the output list
    if email_str != '':
        final_data['Book Output'].append(email_str)
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

worksheet.add_table(0, 0, max_row, max_col - 1, {'columns': column_settings})

worksheet.set_column(0, max_col - 1, 12)
writer.close()  # close and output
