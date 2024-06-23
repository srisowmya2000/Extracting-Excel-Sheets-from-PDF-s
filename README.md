Excel Extraction from Reports

Introduction:
The code has been developed for CSV file extraction from the pdf report.  This is created using Python code which extracts CSV files to the output directory When the directory of the pdf is given. 

Requirements: 
Python IDE:  Vscode, Jupiter notebook, PyCharm.
Pdftk Installation supports – Windows, Mac, Linux. 

Libraries 

Import os -- Provides a portable way to interact with the operating system from Python.
Import subprocess – Allows you to spawn new processes (programs), connect to their input/output/error pipes, and capture their return codes.
Import shutil -- provides high-level functions for working with files and directories, often simplifying common tasks.

Features used from Lib:
os.makedirs(output_directory, exist_ok=True) –  Create new directory if doesn’t exist. If it exists it doesn’t create. 
os.walk(pdf_directory) -- Iterates through all files and subdirectories within the provided directory. 
os.path.join(output_directory, 'temp') -- Combines (joins) path components to create a full path.
subprocess.run -- This line executes the pdftk command with specific arguments to unpack attachments from a PDF.
Root: The path to the current directory being processed.
files: A list of filenames within the current directory.

Steps to execute:  
Open one of your Python IDE. 
Open the code in python. 
Edit the path where the pdf is present and then give the output path. 
Execute the python code


Reference: 
https://docs.python.org/3/library/os.html
https://docs.python.org/3/library/shutil.html
https://python.readthedocs.io/en/latest/library/subprocess.html?highlight=re 

