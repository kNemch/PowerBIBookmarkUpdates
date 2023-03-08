import argparse
import os
import shutil
import re
import calendar
import datetime
import json

from zipfile import ZipFile


"""GLOBAL VARIABLES"""
# These are the default path values. They may be replaced if the user enters custom directory path using CLI
"""Path constants"""
WORKSPACE_PATH:str = None   # workspace directory with all the data
SOURCE_PATH:str = None      # directory with original .pbix files
TEMP_PATH:str = None        # directory for temporary files
RESULTS_PATH:str = None     # directory for the modified .pbix files


"""FUNCTIONS"""
def get_cli_args():
    parser = argparse.ArgumentParser(
        description='''============================
        \rPBI Bookmarks Monthly Update
        \r============================
        \rThe script updates the Year, Month and Quarter values in PBI report's slicers and bookmarks.

        \rBUSINESS LOGIC of the update process:
        \r- Every 15th day from the beginning of the month the default slicers values should be changed in the reports and bookmarks.
        \r- The new default values should be the ones that represent the previous month.

        \rHOW TO USE THE SCRIPT
        \rYou can use the script either with or without the CLI arguments.
        \r- Directory argument (-d --directory):
        \r  * Enter the argument to specify the custom location of the "workspace" directory.
        \r  * Ignore the argument, if the script is in the "workspace" directory.
        \r- Year and Month arguments (-y --year & -m --month)
        \r  * Enter the arguments to specify the custom new values of Year and Month.
        \r    The Year and Month arguments should be passed together.
        \r  * Ignore the arguments, to update the report according to current date.
        ''',
        formatter_class=argparse.RawDescriptionHelpFormatter
    )
    parser.add_argument("-d", "--directory", type=str, help="Directory, which will be the workspace for updates", required=False)
    parser.add_argument("-y", "--year", type=int, help="New value for Year", required=False)
    parser.add_argument("-m", "--month", type=int, help="New value for Month", required=False)
    return parser.parse_args()

def setup_worspace_paths():
    WORKSPACE_PATH = os.getcwd()  # workspace directory with all the data
    SOURCE_PATH = os.path.join(WORKSPACE_PATH, "source")  # directory with original .pbix files
    TEMP_PATH = os.path.join(WORKSPACE_PATH, "temp")      # directory for temporary files
    RESULTS_PATH = os.path.join(WORKSPACE_PATH, "result") # directory for the modified .pbix files


if __name__ == "__main__":
    cli_args = get_cli_args()


    print("Directory: {}".format(cli_args.directory))
    print("Year: {} Month: {}".format(cli_args.year, cli_args.month))