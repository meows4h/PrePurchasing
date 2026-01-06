import os
import pandas as pd
import sys
import configparser

# configure this script using the .ini file
full_config = configparser.ConfigParser()
full_config.read('config.ini')
config = full_config['DEFAULT']  # reading the DEFAULT section
debug = config['Debug']
remove_duplicates = config['RemoveDuplicateBooks']
curr_dir = os.path.dirname(__file__)

if debug == 'False':
    debug = False
if remove_duplicates == 'False':
    remove_duplicates = False

if debug:
    print('Debug texts are active.')

# checking for blank input directories
if config['InputDir'] is not None:
    input_path = os.path.join(curr_dir, config['InputDir'],
                              config['InputFile'])
else:
    input_path = os.path.join(curr_dir, config['InputFile'])

if config['OutputDir'] is not None:
    output_path = os.path.join(curr_dir, config['OutputDir'],
                               config['OutputFile'])
else:
    output_path = os.path.join(curr_dir, config['OutputFile'])

if config['SkippedDir'] is not None:
    skipped_path = os.path.join(curr_dir, config['SkippedDir'],
                                config['SkippedFile'])
else:
    output_path = os.path.join(curr_dir, config['SkippedFile'])


def check_dir(directory: str):
    '''If the output path already exists, checks with the user to make sure
    they do not overwrite their already existing data.'''

    if os.path.exists(directory):
        user_inp = input(f'''File {directory} already exists.
Input "y" to continue & overwrite; "n" to exit.
Input: ''')

        if user_inp == 'y':
            pass
        else:
            sys.exit('\nExiting program.')


def get_sheetname():
    '''Asks the user to input what sheet they wish to process.'''
    user_inp = input('''What sheet are you wanting to process?
i.e. Summer25, Fall25, etc.
Input: ''')
    return user_inp


def get_edition_str(row, row_name):
    '''Takes the current working row and checks it for the
       edition number, creating a string for it.'''

    ed_temp = row[row_name]
    edition_num = str(ed_temp) if pd.isna(ed_temp) is False else None
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


def get_course_arr(row):
    '''Takes the current working row and extracts the course code,
       converting it into a single array for usage.'''

    course_code = ''
    last_char = False

    for char in row['Course Number']:

        # adding any letters that appear
        if char.isalpha():
            course_code += char

        # if the first part just finished (all letters), add a space
        # this ensures there is a space between subject & code i.e. MTH 105Z
        elif char.isnumeric() and not last_char:
            course_code += ' '
            course_code += char
            last_char = True

        # adding remainder of numbers
        elif char.isnumeric():
            course_code += char

        # ensuring that hyphens for course codes are handled properly
        elif char == '-':
            course_code += ' '

    course_arr = course_code.split(' ')
    course_arr = [course_arr[0], course_arr[1]]

    return course_arr


def get_access_types(row):
    '''Takes the current working row and extracts
       all of the copy amounts and access links.'''

    return_row = [[], [], [], [], []]
    if not (pd.isna(row['Ebook Permalink']) or pd.isna(row['Ebook Users'])
            or pd.isna(row['CDL'])):
        return_row[0].append(row['Ebook Permalink'])
        return_row[0].append(row['Ebook Users'])
        return_row[0].append(row['CDL'])

    if not (pd.isna(row['Print Permalink 1'])
            or pd.isna(row['Print 1 Copies'])):
        return_row[1].append(row['Print Permalink 1'])
        return_row[1].append(row['Print 1 Copies'])

    if not (pd.isna(row['Print Permalink 2'])
            or pd.isna(row['Print 2 Copies'])):
        return_row[2].append(row['Print Permalink 2'])
        return_row[2].append(row['Print 2 Copies'])

    if not (pd.isna(row['BNC Permalink'])
            or pd.isna(row['BNC Copies'])):
        return_row[3].append(row['BNC Permalink'])
        return_row[3].append(row['BNC Copies'])

    if not pd.isna(row['Audiobook Permalink']):
        return_row[4].append(row['Audiobook Permalink'])

    return return_row


def get_access_email(book):
    '''This function takes in a single books access data and reads across it
       to create an email with the various links and listing for each title.'''

    email_return = ''
    scanned_appear = False
    physical_appear = False
    # Added ebook tracking vairable
    ebook_appear = False

    # holding all the strings
    # this makes it easier to iterate through them in the proper order
    ebook_list = ''
    scan_list = ''
    audio_list = ''
    # video_list = ''
    print_list = ''
    # dvd_list = ''

    # ebooks
    if book[0]:
        if book[0][2] is False:
            # Tracking ebook to make statement in Format Statements area
            ebook_appear = True
            if (book[0][1] == 'non-perm' or book[0][1] == 'nonperm' or book[0][1] == 'unlimited'):
                ebook_list += f'<li><a href="{book[0][0]}">Ebook</a>: unlimited simultaneous users</li>'
            else:
                # adding some extra cases to prevent breaks when data is formatted weird
                if not pd.isna(book[0][1]):
                    user_str = 'users' if int(book[0][1]) != 1 else 'user'
                else:
                    book[0][1] = '???'
                    user_str = 'users'

                ebook_list += f'<li><a href="{book[0][0]}">Ebook</a>: {book[0][1]} simultaneous {user_str}</li>'

    # CDL
    if book[0]:
        if book[0][2] is True:
            scanned_appear = True
            if (book[0][1] == 'non-perm' or book[0][1] == 'unlimited'):
                scan_list += f'<li><a href="{book[0][0]}">Scanned Book</a>: unlimited simultaneous users</li>'
            else:

                if not pd.isna(book[0][1]):
                    user_str = 'users' if int(book[0][1]) != 1 else 'user'
                else:
                    book[0][1] = '???'
                    user_str = 'users'

                scan_list += f'<li><a href="{book[0][0]}">Scanned Book</a>: {book[0][1]} simultaneous {user_str}</li>'

    # print books
    print_copies = 0
    if book[1]:
        print_copies += int(book[1][1])
        print_link = book[1][0]

    if book[2]:
        print_copies += int(book[2][1])

    if print_copies > 0:
        physical_appear = True
        copy_str = 'copies' if print_copies != 1 else 'copy'
        print_list += f'<li><a href="{print_link}">Print</a>: {print_copies} {copy_str} in Course Reserves</li>'

    # audiobooks
    if book[4]:
        audio_list += f'<li><a href="{book[4][0]}">Audiobook</a> (limited users)</li>'

    access_listings = [ebook_list, scan_list, audio_list, print_list]
    for listing in access_listings:
        if listing != '':
            email_return += listing
    # Returning new variable called ebook_appear
    return email_return, scanned_appear, physical_appear, ebook_appear


def preserve_missed_entry(row, output_set, note):
    '''Formatting error outputs for the extra Excel sheet.'''
    output_set['Course Number'].append(row['Course Number'])
    output_set['Section'].append(row['Section'])
    output_set['Primary Instructor'].append(row['Primary Instructor'])
    output_set['Email Address'].append(row['Email Address'])
    output_set['Title'].append(row['Title'])
    output_set['Ed'].append(row['Ed'])
    output_set['Year'].append(row['Year'])
    output_set['Author'].append(row['Author'])
    output_set['ISBN'].append(row['ISBN'])
    output_set['Publisher'].append(row['Publisher'])
    output_set['Requirement'].append(row['Requirement'])
    output_set['Comments'].append(row['Comments'])
    output_set['Max Enrollment'].append(row['Max Enrollment'])
    output_set['Actual Enrollment'].append(row['Actual Enrollment'])
    output_set['Unable to Purchase'].append(row['Unable to Purchase'])
    output_set['Ebook Permalink'].append(row['Ebook Permalink'])
    output_set['Ebook Users'].append(row['Ebook Users'])
    output_set['Ebook MMS Id'].append(row['Ebook MMS Id'])
    output_set['CDL'].append(row['CDL'])
    output_set['Print Permalink 1'].append(row['Print Permalink 1'])
    output_set['Print 1 MMS Id'].append(row['Print 1 MMS Id'])
    output_set['Print 1 Copies'].append(row['Print 1 Copies'])
    output_set['Print Permalink 2'].append(row['Print Permalink 2'])
    output_set['Print 2 MMS Id'].append(row['Print 2 MMS Id'])
    output_set['Print 2 Copies'].append(row['Print 2 Copies'])
    output_set['BNC Permalink'].append(row['BNC Permalink'])
    output_set['BNC Copies'].append(row['BNC Copies'])
    output_set['Audiobook Permalink'].append(row['Audiobook Permalink'])
    output_set['Audiobook MMS Id'].append(row['Audiobook MMS Id'])
    output_set['Everything in Reading List'].append(row['Everything in Reading List'])
    output_set['Date Emailed'].append(row['Date Emailed'])
    output_set['Email Send Error; Fix and Retry'].append(row['Email Send Error; Fix and Retry'])
    output_set['Notes'].append(row['Notes'])
    output_set['Permalink Sent to Beaver Store'].append(row['Permalink Sent to Beaver Store'])
    output_set['Date Beaver Store Emailed'].append(row['Date Beaver Store Emailed'])
    output_set['Error Note'].append(note)
    return output_set


# getting the sheetname
process_sheet = get_sheetname()

# read the data into the script
data = pd.read_excel(input_path, sheet_name=process_sheet)

# check the output directory
check_dir(output_path)
check_dir(skipped_path)

# set up the outline to grab all the necessary information
result_outline = {'Instructor': [],
                  'Email': [],
                  'Books': [],
                  'Book Output': [],
                  'Courses': []}

# format data output for missing entry excel sheet
preserve_data = {'Error Note': [],
                 'Course Number': [],
                 'Section': [],
                 'Primary Instructor': [],
                 'Email Address': [],
                 'Title': [],
                 'Ed': [],
                 'Year': [],
                 'Author': [],
                 'ISBN': [],
                 'Publisher': [],
                 'Requirement': [],
                 'Comments': [],
                 'Max Enrollment': [],
                 'Actual Enrollment': [],
                 'Unable to Purchase': [],
                 'Ebook Permalink': [],
                 'Ebook Users': [],
                 'Ebook MMS Id': [],
                 'CDL': [],
                 'Print Permalink 1': [],
                 'Print 1 MMS Id': [],
                 'Print 1 Copies': [],
                 'Print Permalink 2': [],
                 'Print 2 MMS Id': [],
                 'Print 2 Copies': [],
                 'BNC Permalink': [],
                 'BNC Copies': [],
                 'Audiobook Permalink': [],
                 'Audiobook MMS Id': [],
                 'Everything in Reading List': [],
                 'Date Emailed': [],
                 'Email Send Error; Fix and Retry': [],
                 'Notes': [],
                 'Permalink Sent to Beaver Store': [],
                 'Date Beaver Store Emailed': []}

# format data output for output excel sheet
final_data = {'First Name': [],
              'Last Name': [],
              'Email': [],
              'Book Output': []}


# looking across the whole input spreadsheet
for idx, row in data.iterrows():

    # checking if the email has already been sent
    if not pd.isna(row['Date Emailed']) and pd.isna(row['Email Send Error; Fix and Retry']):
        if debug:
            print(f'Skipping... Already sent ; Inst: {row['Primary Instructor']}, Book: {row['Title']}')
        continue

    # checking reading list
    if row['Everything in Reading List'] is False or pd.isna(row['Everything in Reading List']):
        if debug:
            print(f'Skipping... Not in reading list ; Inst: {row['Primary Instructor']}, Book: {row['Title']}')
        continue

    instructor = str(row['Primary Instructor']).strip()
    email = str(row['Email Address']).strip()

    if instructor == 'STAFF':
        if debug: print(f'Skipping... STAFF listing ; Inst: {row['Primary Instructor']}, Book: {row['Title']}')
        continue

    if instructor == '0' or email == '0' or instructor == '#N/A' or email == '#N/A' or pd.isna(row['Primary Instructor']) or pd.isna(row['Email Address']):
        print(f'Skipping... Missing critical info ; Inst: {row['Primary Instructor']}, Book: {row['Title']}')
        preserve_missed_entry(row, preserve_data, 'Instructor Information')
        continue

    # grabbing the edition number, course code string, and access types/links
    edition_num = get_edition_str(row, 'Ed')
    course_arr = get_course_arr(row)
    access_type = get_access_types(row)

    # access_type input validation
    check_num = 0
    for access in access_type:
        if access:
            check_num += 1
    if not check_num:
        print(f'Skipping... Missing access information entirely ; Inst: {row['Primary Instructor']}, Book: {row['Title']}')
        preserve_missed_entry(row, preserve_data, 'Access Information')
        continue

    # compile baseline information + remove blank space
    title = row['Title'].title()
    title = title.strip()
    author = row['Author'].title()
    author = author.strip()
    book_info = [title, edition_num, author, access_type, course_arr, row['Year']]

    # add new instructor to the outline
    if instructor not in result_outline['Instructor']:
        result_outline['Instructor'].append(instructor)
        result_outline['Email'].append(email)
        result_outline['Books'].append([book_info])
        result_outline['Courses'].append([book_info[4]])

    # add any new information for existing instructors
    else:
        person_idx = result_outline['Instructor'].index(instructor)
        result_outline['Books'][person_idx].append(book_info)

        if book_info[4] not in result_outline['Courses'][person_idx]:
            result_outline['Courses'][person_idx].append(book_info[4])


# taking the result outline and finalizing it into an output for email
# for each instructor found
for instructor in result_outline['Instructor']:

    first_name = ''
    last_name = ''

    name_arr = instructor.split(', ')
    if len(name_arr) == 1:
        name_arr = instructor.split(' ')
        last_name = name_arr[len(name_arr) - 1].strip()
        for piece in name_arr:
            if piece == last_name:
                break
            first_name += f' {piece}'
        first_name = first_name.strip()
    else:
        first_name = name_arr[1].strip()
        last_name = name_arr[0].strip()

    idx = result_outline['Instructor'].index(instructor)

    final_data['First Name'].append(first_name)
    final_data['Last Name'].append(last_name)
    final_data['Email'].append(result_outline['Email'][idx])

    email_str = ''
    scanned_appear = False
    physical_appear = False
    ebook_appear = False

    # for each course the instructor is under
    for k, course in enumerate(result_outline['Courses'][idx]):
        if k != 0:
            email_str += '<br>'

        email_str += f'<b>{course[0]} {course[1]}</b><br>Students can find all library materials by <a href="https://search.library.oregonstate.edu/discovery/search?query=any,contains,{course[0]}%20{course[1]}&tab=CourseReserves&search_scope=CourseReserves&vid=01ALLIANCE_OSU:OSU&lang=en&offset=0">searching course reserves for {course[0]} {course[1]}</a>.<br><ul>'

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

            if book_course != course:
                continue

            # removing these repeated tags
            cleaner_list = [' (Cei)', '(Loose-Leaf)', 'Loose-Leaf', 'Ebook - Lifetime Duration', 'Ebook(5 Yr Access)', 'Ebook (Lifetime)', 'Ebook (180 days)', 'Ebook (150 days)', 'Ebook (120 days)', 'Ebook - Lifetime Access', 'Ebook -Lifetime Access', 'Ebook - Lifetime', 'Ebook - 180Days', 'Etext W/Connect Access Code', '[Qr]', '[Nbs]', '(Cei)', '(Ll)', 'W/1 Term Access Code Pkg', 'W/1 Year Access Code Pkg', 'W/2 Year Access Code Pkg', '1 Term Access Code', '1 Year Access Code', '2 Year Access Code', 'Ebook']
            if remove_duplicates:
                for phrase in cleaner_list:
                    book_title = book_title.replace(phrase.title(), '')

            # NOTE : Maybe fix authors with Mc- in the start of their name?
            book_author = book_author.replace('(Digital)', '')
            book_author = book_author.replace('(2)', '')

            # removing trailing whitespace
            book_title = book_title.strip()
            book_author = book_author.strip()

            # removing odd capitalization after apostrophes
            last_char = ''
            for i, character in enumerate(book_title):
                if last_char == "'":
                    temp_list = list(book_title)
                    temp_list[i] = book_title[i].lower()
                    book_title = ''.join(temp_list)
                last_char = character

            if ' Of ' in book_title:
                book_title = book_title.replace(' Of ', ' of ')

            # skip if duplicate
            if book_title in used_books and remove_duplicates:
                continue

            # adding in the book listing header
            email_str += f'<li><em>{book_title}</em>, {book_author}'
            email_str += f', {book_edition}' if book_edition is not None else ''
            email_str += f', {book_year}</li>' if not pd.isna(book_year) else '</li>'

            # check access types and add new email pieces to the main email string
            access_email, scanned_appear, physical_appear, ebook_appear = get_access_email(book_access)
            if access_email != '':
                email_str += f'<ul>{access_email}</ul>'

            used_books.append(book_title)

        email_str += '</ul>'

    # checking whether or not these specific cases
    # have appeared for this professor
    # Scanned was updated with new language according to Pre-Purchasing Email Notification Doc
    # Ebook language was also added according to same doc mentioned above
    if scanned_appear:
        email_str += '<br>Scanned books are first come, first serve, for one hour at a time and use a waitlist. There is no limit to the number of renewals if no one is in the waitlist. If you would like to increase the number of simultaneous users, you may provide additional print copies for us to sequester. For every print copy we sequester, we can allow one online user in accordance with <a href="https://www.controlleddigitallending.org/">controlled digital lending</a> principles.'
    if scanned_appear and physical_appear:
        email_str += '<br>'
    if physical_appear:
        email_str += '<br>Physical copies are available for checkout at the Borrowing & Information desk for three hours at a time.'
    if (scanned_appear and ebook_appear) or (physical_appear and ebook_appear):
        email_str += '<br>'
    if ebook_appear:
        email_str += '<br>When purchasing ebooks, we prioritize unlimited-user licenses, but these are not always available.'

    final_data['Book Output'].append(email_str)


def write_to_excel(directory, export_data, sheetname):
    '''Exports given data to given directory with given sheetname.'''
    dataframe = pd.DataFrame(data=export_data)
    writer = pd.ExcelWriter(directory, engine='xlsxwriter')
    dataframe.to_excel(writer, sheet_name=sheetname, startrow=1,
                       header=False, index=False)
    worksheet = writer.sheets[sheetname]
    (max_row, max_col) = dataframe.shape
    column_settings = []
    for header in dataframe.columns:
        column_settings.append({'header': header})
    worksheet.add_table(0, 0, max_row, max_col - 1,
                        {'columns': column_settings})
    worksheet.set_column(0, max_col - 1, 12)
    writer.close()


write_to_excel(output_path, final_data, 'Email List')
if len(preserve_data['Title']):
    write_to_excel(skipped_path, preserve_data, process_sheet)
else:
    print('No errors found outside of expected values.')
