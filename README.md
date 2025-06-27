# Pre-Purchasing Email Generator
This is supposed to take the prepurchasing sheet and convert the necessary listings into a single email per professor, divided per course.

## Getting Started
If you do not know how to use a command line (cmd.exe) feel free to just run the `install.bat` file, it runs the pip command for you.

The only python package/library this script needs is pandas. Regardless, feel free to use the following:
```pip install requirements.txt```

By default, the config is configured to intake an `input.xlsx` file and output an `output.xlsx` file. This can be changed in the `config.ini` file.

## General Script Usage
To run the script, you can either use `run.bat` or run it using python from a command line.

If you already have a matching output file in the same output directory, it will ask to confirm that you wish to overwrite it.

It will then ask if you would like to remove duplicates (all the entires with Ebook Unlimited Access, etc), you likely want to say yes.

After this, it will run, if it has debug on, it will spit out a ton of entries it skips over due to not having information or not being included.

Finally, you will take this output and move it into the SharePoint folder for use in the flow. Running it directly from the flow will send the formatted emails.

## Full Step by Step Guide
### Script Portion
1. Download the full or part of the Pre-Purchasing sheet by going to File -> Create a Copy -> Download a Copy.
2. Rename the file to `input.xlsx` (or whichever name you set in `config.ini`).
3. Move the file to the directory of the python script (or to where you set the input directory).
4. If you have not already, run `install.bat` (or install pandas via pip).
5. Run `run.bat` (or `main.py` via the command line). (Note: Running this for the first time may have some lag, be patient! If you need to restart it, press Ctrl + C once or twice to stop the program)
6. You may be asked if you wish to overwrite the output file, hit yes or no depending on what you'd like.
7. You will be asked if you wish to remove duplicates, enter 'y' to indicate yes.
8. The script will run, if debug is enabled in `config.ini` it will print all skipped entries.
9. There will now be a file named `output.xlsx` (or whichever name you set in `config.ini`).

### Power Automate Portion
1. Take this output file and feel free to rename it to what seems right (i.e. Summer25 Emails)
2. Upload the file to the appropriate folder of the SharePoint.
3. Go to the Power Automate Pre-Purchasing Flow, click on edit, and click on "List rows present in a table". (Find this Flow at https://make.powerautomate.com/environments/Default-ce6d05e1-3c5e-4d62-87a8-4c4a2713c113/flows/shared/286b5cad-8ed1-48d6-9358-5ab72464b029/details?v3=false)
4. Here you will select "Group - OSU Libraries and Press" and "Documents" for the first two boxes if they are not already populated, if they are, continue.
5. The third box is to the directory of where the file is, feel free to click through until you find it (i.e. /LEAD/eTextbooks project/Summer25 Emails.xlsx).
6. Finally, click on the last box and make sure "Table1" is selected.
7. For the email formatting, go into the "Apply to each" tab and into the "Send an email from a shared inbox", make sure the Subject is correct, as well as the email format.
8. Hit save and then feel free to run the Flow, it will send out all the emails listed in the spreadsheet generated from the script, all of them will fill the Drafts before being sent out in batches.
