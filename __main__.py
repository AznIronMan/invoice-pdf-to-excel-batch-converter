import json
import os
import sys


from pathlib import Path


## ADJUSTABLE VARIABLES ##

# FORCE_DEBUG should be a boolean (True or False, no quotes)
# Note: Environment variables can only store strings. Convert this to string if it has to be an environment variable.
FORCE_DEBUG = True

# HEADER_FILL should be a string array
# The first value is the background color, the second value is the font color, and the third value is the fill type.
# EDIT THIS AS NEEDED
HEADER_FILL = ['4CAF50', '4CAF50', 'solid']

# LOG_DIR should be a string
# This would be in the root directory where this __main__.py file is located.
LOG_DIR = 'logs'

# PROCESSED_DIR is used to set the processed directory that will be created in the found PDF directory.
# This is where the processed Excel files will be stored.
PROCESSED_DIR = 'processed'

# CELL_PHONE should be a string array
# EDIT THIS AS NEEDED
CELL_PHONE = ['Cell', 'Mobile', 'iPhone']


# MAIN_PHONE should be a string array
# EDIT THIS AS NEEDED
MAIN_PHONE = ['Tel', 'Main', 'Home', 'Office', 'Phone', 'Telephone']


## DO NOT CHANGE ANYTHING BELOW THIS LINE ##

# This sets the environment variables for the program.
os.environ['CELL_PHONE'] = json.dumps(CELL_PHONE)
os.environ['FORCE_DEBUG'] = str(FORCE_DEBUG)
os.environ['HEADER_FILL'] = json.dumps(HEADER_FILL)
os.environ['MAIN_PHONE'] = json.dumps(MAIN_PHONE)
os.environ['LOG_DIR'] = str(LOG_DIR)
os.environ['PROCESSED_DIR'] = str(PROCESSED_DIR)
os.environ['TK_SILENCE_DEPRECATION'] = '1'

# This sets the root directory to the parent directory of this __main__.py file.
current_path = Path(__file__).resolve()
parent_path = current_path.parent.parent
sys.path.append(str(parent_path))

# This is the auto-start for the program.
if __name__ == "__main__":
    from core.process import batch_convert
    from core.filer import select_folder
    os.system('cls' if os.name == 'nt' else 'clear')
    batch_convert(select_folder())
