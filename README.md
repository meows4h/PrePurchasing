# Pre-Purchasing Email Generator
This is supposed to take the prepurchasing sheet and convert the necessary listings into a single email per professor, divided per course.

## Getting Started
If you do not know how to use a command line (cmd.exe) feel free to just run the `install.bat` file, it runs the pip command for you.

The only python package/library this script needs is pandas. Regardless, feel free to use the following:
```pip install requirements.txt```

By default, the config is configured to intake an `input.xlsx` file and output an `output.xlsx` file. This can be changed in the `config.ini` file.

## Using the Script
To run the script, you can either use `run.bat` or run it using python from a command line.
If you already have a matching output file in the same output directory, it will ask to confirm that you wish to overwrite it.
It will then ask if you would like to remove duplicates (all the entires with Ebook Unlimited Access, etc), you likely want to say yes.
After this, it will run, if it has debug on, it will spit out a ton of entries it skips over due to not having information or not being included.

Finally, you will take this output and move it into the SharePoint folder for use in the flow. Running it directly from the flow will send the formatted emails.

This is still a big work in progress and this README likely needs some more specifics and cleanup.
