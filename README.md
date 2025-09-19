# Textbook Resources Email Generator
This program contains two main scripts, one for each spreadsheet input type. The main part of this program is for the Pre-Purchasing sheet, while the other additional piece is for the Requests/Purchasing sheet.

## Getting Started
To install the required Python libraries, use the `install.bat` file, it will automatically run the command:
```pip install -r requirements.txt```

Subsequently after installing the libraries, to run the script, use the `run.bat` file. This will run the command for the primary script:
```python main.py```

If it were for the requests script, the `run.bat` file runs:
```python requests.py```

## Main Pre-Purchasing Script Usage
Within the main folder, to run the script, use either `run.bat` or open Command Prompt, navigate to the folder directory using the `cd` (change directory) command, and run `python main.py`.

The first question it will ask is for the name of the sheet being processed, often this will look like something like "Fall25" or "Summer25", etc. Input which sheet is looking to be processed.

Next, it may state that the output file already exists, as well as the errors/additional output file already existing. If you're sure you're fine with overwriting each of these, input `y` in each case as prompted.

After that last question, it should run the remainder of the script and output the relevant file.

If it crashes or has any serious issues following this, please reach out to Rox Beecher to correct it.

## Requests Script Usage
Within the requests subfolder, to run the script, use either `run.bat` or open Command Prompt, navigate to the folder directory using the `cd` (change directory) command, and run `python requests.py`.

It may state that the output file already exists. If you're sure you're fine with overwriting it, input `y` in as prompted.

After the prompt, it should run the remainder of the script and output the relevant file. It may list a considerable amount of statements that show `WARNING` and `ERROR`.

`WARNING` indicates something minor is missing, but not integral to the output, merely warning you about it missing. 

`ERROR` indicates something major is missing, that is integral to the output. In each case of an error, it is skipping over the listed row or book as it cannot write it to an email.

If it crashes or has any serious issues following this, please reach out to Rox Beecher to correct it.

## Power Automate
This is the step that directly ties in and directly sends out the emails.

After running the script and getting an output file, as long as the output file is located within a OneDrive / Sharepoint area, the Flow should be able to link to them.

Navigate to the necessary Flow for the output you are processing.

To send out the current batch, click on `Edit`, click on the `List rows present in a table` node which should open a menu on the left hand side. Go ahead and click on the `File` input and ensure it is set to the new output that was just generated. After this, change the `Table` input below it to the only selectable table option (usually `Table1`). After setting that, it should be good to go.

If you'd like to modify the exact template, click on the drop down arrow on `Apply to each` to reveal the `Send an email from a shared mailbox (V2)` node. It will reveal the menu on the left hand side, similar to an email interface. The `Book Output` variable is going to be all the textbooks and automatically generated associated text, everything else around the text can be modified to the necessary fit of the script.

After each piece is done, save it, run it, and confirm that each email is being sent from the proper Outlook inbox.

If any further instruction is needed, reach out to Rox Beecher.

## File/Script Configuration
There are a couple options inside the `config.ini` file. In order to edit these options, feel free to open the file within a program like Notepad. Each of these values should have a comment next to them within the file.

If in doubt, don't touch these values, as they can have incorrect values that changes / crashes the script.

### Main Pre-Purchasing Config
- `Debug`: `True` or `False` (Enables output that shows extra checks)
- `RemoveDuplicateBooks`: `True` or `False` (Enables removing duplicate titles under the same professor)
- `InputDir`: Directory; i.e. C:/Users/<user>/Downloads/<file>, if left blank, looks in same folder (What file to search for the input file within)
- `InputFile`: File Name; i.e. Data AY25-26.xlsx (Specific name of file to input)
- `OutputDir`: Directory; if left blank, outputs to same folder
- `OutputFile`: File Name; i.e. output.xlsx (Specific name of file to output)

### Requests Config
- `Debug`: `True` or `False` (Enables output that shows extra checks)
- `Feedback`: `True` or `False` (Enables output to show flagged missing information)
- `SkipLinks`: `True` or `False` (Enables skipping over entries with missing `https://search...` links)
- `RemoveDuplicateBooks`: `True` or `False` (Enables removing duplicate titles under the same professor)
- `InputDir`: Directory; i.e. C:/Users/<user>/Downloads/<file>, if left blank, looks in same folder (What file to search for the input file within)
- `InputFile`: File Name; i.e. Data AY25-26.xlsx (Specific name of file to input)
- `SheetName`: Excel Sheetname; i.e. CourseReserves-VAL-Active-25-26 (Within Excel, name of sheet)
- `OutputDir`: Directory; if left blank, outputs to same folder
- `OutputFile`: File Name; i.e. output.xlsx (Specific name of file to output)
