import argparse
import os
import shutil
import re
import calendar
import datetime
import json

from zipfile import ZipFile


"""GLOBAL VARIABLES"""
"""Path constants"""
# The values for these global path variables are assigned in the setup_worspace_paths() function, depending on the CLI user input
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


def setup_worspace_paths(cli_workspace_path:str):
    """Assigning the paths to the global variables"""
    global WORKSPACE_PATH, SOURCE_PATH, TEMP_PATH, RESULTS_PATH
    WORKSPACE_PATH = cli_workspace_path if cli_workspace_path else os.getcwd()  # workspace directory with all the data
    SOURCE_PATH = os.path.join(WORKSPACE_PATH, "source")  # directory with original .pbix files
    TEMP_PATH = os.path.join(WORKSPACE_PATH, "temp")      # directory for temporary files
    RESULTS_PATH = os.path.join(WORKSPACE_PATH, "result") # directory for the modified .pbix files


def get_the_list_of_pbix_files():
    for folderName, _, filenames in os.walk(WORKSPACE_PATH):
        for filename in filenames:
            if filename[-5:] == ".pbix" and not folderName.__contains__():
                # create complete filepath of file in directory
                temp_dir_file_path = os.path.join(folderName, filename)
                # path for the archive hierarchy
                archive_path = temp_dir_file_path[len(temp_dir_path)+1:]
                # selecting the compression mode depending on archive file type
                # writing the file to result archive

                print(folderName, subfolders, filename, temp_dir_file_path, archive_path)


def create_workspace_hierarchy():
    """Checking if directories exist"""
    if not os.path.exists(RESULTS_PATH):
        os.mkdir(RESULTS_PATH)
    if not os.path.exists(TEMP_PATH):
        os.mkdir(TEMP_PATH)
    if not os.path.exists(SOURCE_PATH):
        os.mkdir(SOURCE_PATH)
        print("Please, place all the .pbix files into the 'source' directory, and then rerun the script")
    else:
        print("The directories structure is OK")
    





if __name__ == "__main__":
    cli_args = get_cli_args()
    print("Workspace Directory: {}".format(cli_args.directory))
    print("New values for the reports: Year - {}, Month - {}".format(cli_args.year, cli_args.month))

    setup_worspace_paths(cli_args.directory)
    create_workspace_hierarchy()
    



    

    print(WORKSPACE_PATH, SOURCE_PATH, TEMP_PATH, RESULTS_PATH, sep="\n")

    