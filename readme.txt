========
PPTX2TXT
========

OVERVIEW
========

PPTX2TXT is a simple Python script that extracts all text from a .pptx file and saves it as a .txt file (with the same name as the .pptx file) in the same directory as the .pptx file.

The .txt file contains separators to indicate which slide content has come from, as well as in-line labels indicating whether the text comes from a user-placed shape or table, or from a layout element in the slide template (e.g., a footer). Content from the presentation's overall slide master is also included at the end.

Microsoft PowerPoint no longer includes a feature for comparing presentations to highlight differences. With PPTX2TXT, you can extract the text from two different versions of a presentation, and then use a diff tool (e.g., WinMerge on PC) to identify differences.


REQUIREMENTS
============

- Python 3.13+ 
- Library: python-pptx

If you already have Python 3.13+ installed on your computer, you can simply add the python-pptx library to your system. However, it is better to create a Python virtual environment (or 'venv') specifically for PPTX2TXT, install the python-pptx library in the venv, and run the main PPTX2TXT script from that venv.

To save you time, an interactive Python script has been provided that will automatically create a venv, adding the required libraries and solution scripts. 


INSTALLATION
============

1. Ensure Python 3.13+ is installed on your system (download from https://www.python.org/).

2. Ensure Python is added to your system's PATH (this is usually done automatically during installation but may require a system restart).

3. In a cmd/terminal window, navigate to the location of this readme.txt file.

4. In the cmd/terminal window, use the following command:

	PC:	python install.py
	MacOS:	Python3 install.py

5. Follow the prompts in the cmd/terminal window, and use file explorer/finder to choose the directory where you would like the venv to be created.

6. When you see the message "PPTX2TXT installation complete" in the cmd/terminal window, PPTX2TXT is ready to use.


USAGE
=====

1. In a cmd/terminal window, navigate to the venv directory.

2. In the cmd/terminal widow, use the following command to ACTIVATE the venv:

	PC: 	Scripts\activate
	MacOS:	source bin/activate

When the venv is activated your cmd/terminal prompt will look similar to this:

	PC:	(venv_ppt2txt) C:\venv_ppt2txt>
	MacOS:	(venv_ppt2txt) you@your-mac venv_ppt2txt %

3. In the cmd/terminal widow, use the following command to start PPTX2TXT:

	PC:	python pptx2txt.py
	MacOS:	Python3 pptx2txt.py

4. Follow the prompts in the cmd/terminal window, and use file explorer/finder to choose the .pptx file you want to extract text from.

5. When you see the message "Output .txt file saved successfully" in the cmd/terminal window, the .txt file will be ready for you to use, and located in the same directory as the .pptx file you selected in step 4. Repeat steps 3-4 to extract text from additional .pptx files.

NOTE: If you have previously generated a .txt file for a .pptx file, PPTX2TXT will tell you the file already exists and tell you to delete it and try again. If this happens, repeat steps 3-4.

6. When you are ready to stop using PPTX2TXT, deactivate the venv by using the following command in the cmd/terminal window:

	PC: deactivate
	MacOS: deactivate

When the venv is deactivated your cmd/terminal prompt will look similar to this:

	PC:	C:\venv_ppt2txt>
	MacOS:	you@your-mac venv_ppt2txt %


WHAT'S IN THE TEXT FILE
=======================

The .txt file is divided into sections for each slide. Within a slide's section, text extracted from each shape is presented in its own line or paragraph. A label at the start of each line or paragraph indicates the type of shape from which the text was extracted:

	[SHAPE] - A user-placed shape (e.g., text box, rectangle)

	[SHAPE TABLE] - A user-placed table

	[LAYOUT] - A layout element from the slide's template (e.g., heading, footer)

	[LAYOUT TABLE] - A layout element that is a table

	[NOTES] - Text from the notes pane of a slide

	[MASTER] - An element from the presentation's slide master

NOTE: For tables, cell texts are tab-separated so you can copy/paste the whole block into Excel. However, within an individual cell's text, all carriage returns have been replaced with triple spaces (allowing you to use search-and-replace to restore the carriage returns).


LIMITATIONS
===========

- PPTX2TXT cannot extract text that is part of a graphic (e.g., text within a logo).

- PPTX2TXT cannot yet extract the text from comments and replies.

- Where a slide contains an empty layout element (e.g., an unused footer), the .txt file may contain 'empty' entries like this:

	[LAYOUT]


CREDITS
=======

Version: 1.1.0 | 26 May 2025
This script was created by Oxzeon Limited (oxzeon.com) with assistance from ChatGPT.
No copyright is asserted - please use/modify/redistribute as you wish.
CAUTION: Use entirely at your own risk.
Contact: alan@oxzeon.com
