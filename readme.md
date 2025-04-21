# Application Name

Folio-to_workday Invoice Mapper

# Purpose

the library currently keeps track of financial data about orders in our Folio ILS (https://folio.org/).  However, data about orders also needs to be shared with the university purchasing department.  They use our campus Workday system (https://www.workday.com/).

Having staff input financial data by hand into both systems is inefficiant and labor intensive.  This simple command-line script is designed to remove some of that burden by mapping a CSV export file from Folio representing recent orders to a different CSV structure that Workday can import.

# Environment Setup/Local Setup

To run this script, the user will need python3 installed on their local computer.  There are plenty of web tutorials that can show you how to do this.  on a mac, the easiest way to do it is to install the xcode command-line tools (NOT the entire Xcode application-that this huge and will take hours to download and install).

# Usage

Once python3 is installed, move the processFile.py file to your local computer.  If you need help installing python3 or hgetting the file, contact the application maintainer, below.

Before you start, be sure you've moved the Folio invoice export file to be processed to the same directory where the processFile script resides.

To use the script, first fire up the windows command prompt:

Click on the "start" menu, then type "cmd" into the blank and hit return.  On a mac, this is under "utilities" in your application folder.

A command prompt window should appear.  To execute the script, you'll need to move your command prompt to the directory where the script and the ingest file are located:

At the prompt, type:  

	cd /path/to/file

Once in the directory, you'll need to start the script.  To do this, you'll need three things:

The name of the script itself (processFile.py)
The name of the file to be transformed
The name of the framework needed to use to interpret and run the script (python)

The command will look like this:

	python3 processFile.py name_of_file_to_transform

Again, it's important that the processFile.py script and the file to be transformed are in the same directory.  If they are not, the script will fail to find the file and shut down.

The script includes a number of file integrity checks:

It checks that all required columns are present and correctly named
It checks for any duplicate required columns 

It checks for proper CSV syntax (commas where they need to be, lines properly terminated, etc.)

If any of these checks fail, the script will shut down and display an error message indicating what the problem is and where.  You'll need to correct it before you can run the script.  If you can't make sense of a specific error message, let me know-it may be an issue I didn't anticipate.  You can also request that I reword error messages to make them clearer-I welcome this kind of feedback!

If it's successful, the script will tell you so and deposit a new file in the directory called "workday.csv".  You can open this is any program that can interpret csv files to verify that it's correct.

If you run the script when there is a "workday.csv" file already in the directory, it will overwrite the existing file with a new one.

# Maintainer

Kyle Felker (felkerk@gvsu.edu)




