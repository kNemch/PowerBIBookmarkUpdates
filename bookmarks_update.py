import argparse
import os
import shutil
import re
import calendar
import datetime
import json
import zipfile

from zipfile import ZipFile


"""GLOBAL VARIABLES"""
"""Path constants"""
# The values for these global path variables are assigned in the setup_work_dir_paths() function, depending on the CLI user input
WORK_DIR_PATH:str = None    # workspace directory with all the data
SOURCE_DIR_PATH:str = None  # directory with original .pbix files - CURRENTLY NOT IN USE
TEMP_DIR_PATH:str = None    # directory for temporary files
RESULTS_DIR_PATH:str = None # directory for the modified .pbix files


"""FUNCTIONS"""
# CLI
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
    parser.add_argument("-p", "--previous-year", type=int, help="Old value for Year", required=False)
    return parser.parse_args()


# WORKING DIRECTORY - PATHS, DIRECTORIES AND FILES
def setup_work_dir_paths(cli_work_dir_path:str):
    """Assigning the paths to the global variables"""
    global WORK_DIR_PATH, SOURCE_DIR_PATH, TEMP_DIR_PATH, RESULTS_DIR_PATH
    WORK_DIR_PATH = cli_work_dir_path if cli_work_dir_path else os.getcwd()  # workspace directory with all the data
    TEMP_DIR_PATH = os.path.join(WORK_DIR_PATH, "#TEMP")      # directory for temporary files
    RESULTS_DIR_PATH = os.path.join(WORK_DIR_PATH, "#RESULTS") # directory for the modified .pbix files
    # SOURCE_PATH = os.path.join(WORK_DIR_PATH, "source")  # directory with original .pbix files


def get_pbix_workspaces_and_files_pairs() -> list[list[str, str]]:
    pbix_paths_and_files = []
    for directory_path, _, filenames in os.walk(WORK_DIR_PATH):
        for filename in filenames:
            if filename[-5:] == ".pbix" and RESULTS_DIR_PATH not in directory_path and TEMP_DIR_PATH not in directory_path:
                directory_with_pbix_path = directory_path[len(WORK_DIR_PATH) + 1:]
                pbix_paths_and_files.append([directory_with_pbix_path, filename])
    return pbix_paths_and_files


def create_directories_hierarchy(pbi_workspaces):
    print("\nCREATING THE DIRECTORIES HIERARCHY")

    directories_to_create = [RESULTS_DIR_PATH, TEMP_DIR_PATH]

    for tech_dir in directories_to_create:
        if not os.path.exists(tech_dir):
            os.mkdir(tech_dir)
            print("Created tech directory: {}".format(tech_dir))
        else:
            print("Tech directory exists: {}".format(tech_dir))
        for pbi_ws_name in pbi_workspaces:
            pbi_ws_subdirectory_path = os.path.join(tech_dir, pbi_ws_name)
            if not os.path.exists(pbi_ws_subdirectory_path):
                os.mkdir(pbi_ws_subdirectory_path)
                print("Created directory for {} workspace: {}".format(pbi_ws_name, pbi_ws_subdirectory_path))
            else:
                print("Workspace {} directory exists: {}".format(pbi_ws_name, pbi_ws_subdirectory_path))


# ARCHIVE OPERATIONS
def unzip_pbix(ws_subdirectory, pbix_filename):
    # paths for the original .pbix file (archive) and temporary folder, to which the file will be unarchived
    source_pbix_file_path = os.path.join(WORK_DIR_PATH, ws_subdirectory, pbix_filename)
    unarchived_temp_files_dir_path = os.path.join(TEMP_DIR_PATH, ws_subdirectory, pbix_filename[:-5])

    # opening the zip file in READ mode
    with ZipFile(source_pbix_file_path, 'r') as source_archive:
        source_archive.extractall(path=unarchived_temp_files_dir_path)
        print('Extracting all the files from', source_pbix_file_path, "to", unarchived_temp_files_dir_path)


def zip_pbix(ws_subdirectory, pbix_filename):
    result_pbix_file_path = os.path.join(RESULTS_DIR_PATH, ws_subdirectory, pbix_filename)
    unarchived_temp_files_dir_path = os.path.join(TEMP_DIR_PATH, ws_subdirectory, pbix_filename[:-5])

    with ZipFile(result_pbix_file_path, mode="w") as result_archive:
        for directory_path, _, filenames in os.walk(unarchived_temp_files_dir_path):
            for filename in filenames:
                # create complete filepath of file in directory
                file_location_in_fs = os.path.join(directory_path, filename)
                # path for the archive hierarchy
                file_location_in_archive = file_location_in_fs[len(unarchived_temp_files_dir_path)+1:]
                # selecting the compression mode depending on archive file type
                file_compression_mode = zipfile.ZIP_STORED if filename == "DataModel" else zipfile.ZIP_DEFLATED
                # writing the file to result archive
                result_archive.write(
                    file_location_in_fs, 
                    arcname=file_location_in_archive, 
                    compress_type=file_compression_mode)
    
    print('Writing all the temporary files to', result_pbix_file_path, "from", unarchived_temp_files_dir_path)


# PBIX MODIFICATION
def remove_security_bindings_data(path_to_unarchived_pbix):
    security_bindings_file_path = os.path.join(path_to_unarchived_pbix, "SecurityBindings")
    if os.path.exists(security_bindings_file_path):
        os.remove(security_bindings_file_path)

    content_types_file_path = os.path.join(path_to_unarchived_pbix, "[Content_Types].xml")

    with open(content_types_file_path, "r") as content_types_file:
        xml = content_types_file.read()

    updated_xml = xml.replace('<Override PartName="/SecurityBindings" ContentType="" />', "")
    with open(content_types_file_path, "w") as content_types_file:
        content_types_file.write(updated_xml)


def get_patterns_and_replacements(cli_arg_year, cli_arg_month):
    if not cli_arg_month and not cli_arg_year:
        current_date = datetime.date.today()
        period_for_new_values = current_date - datetime.timedelta(days=current_date.day)
    else:
        period_for_new_values = datetime.date(cli_arg_year, cli_arg_month, 1)
    
    previous_month = period_for_new_values - datetime.timedelta(days=period_for_new_values.day)

    new_value_year = period_for_new_values.year
    new_value_month = period_for_new_values.month
    new_value_month_abbr = calendar.month_abbr[new_value_month]
    new_value_quarter = (new_value_month + 2) // 3

    period_replacements = [
        # pattern, new value
        ('(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)', new_value_month_abbr),  # month eng
        ('\d{1,2}月', str(new_value_month) + "月"),   # month chn
        ('Q [1-4]', 'Q ' + str(new_value_quarter)),   # quarter eng (with space)
        ('Q[1-4]', 'Q' + str(new_value_quarter)),     # quarter eng (without space)
        ('[1-4]季度', str(new_value_quarter) + "季度"),  # quarter chn
        (str(previous_month.year) + "L", str(new_value_year) + "L"),  # year
        (str(previous_month.year), str(new_value_year))  # year
    ]
    return period_replacements


def replace_period(text:str, period_pattern:str, period_new_value:str):
    search_pattern = r'\\"Value\\":\\"\'{}\'\\"'.format(period_pattern)
    new_value = r'\\"Value\\":\\"' + "'{}'".format(period_new_value) + '\\"'
    return re.sub(search_pattern, new_value, text)



"""OUTPUT FUNCTIONS (STDOUT)"""
def print_cli_input(cli_args):
    print("""
    \r========================
    \rMONTHLY BOOKMARKS UPDATE
    \r========================
    \r\nCLI - USER INPUT
    \rWorkspace Directory: {}
    \rNew values for the reports: 
    Year - {}
    Month - {}"""
          .format(cli_args.directory, cli_args.year, cli_args.month))


def print_patterns(patterns_and_new_values):
    print("\nPATTERNS")
    for period_pattern, period_new_value in period_replacements:
        print("{}  ->  {}".format(period_pattern, period_new_value))



"""MAIN FUNCTION"""
if __name__ == "__main__":

    # Reading user CLI input
    cli_args = get_cli_args()
    print_cli_input(cli_args)

    # Generating patterns and new values
    period_replacements = get_patterns_and_replacements(cli_args.year, cli_args.month)
    print_patterns(period_replacements)

    # Creating directories, getting the list of PBIX files with paths to them
    setup_work_dir_paths(cli_args.directory)
    pbix_workspaces_and_files = get_pbix_workspaces_and_files_pairs()
    pbix_workpaces = set([pbix[0] for pbix in pbix_workspaces_and_files])
    create_directories_hierarchy(pbix_workpaces)

    # Processing of the .pbix files
    print("\nFILES PROCESSING")
    for subdirectory, filename in pbix_workspaces_and_files:
        print("WORKSPACE", subdirectory, "FILE", filename)
        unzip_pbix(subdirectory, filename)
        temp_pbix_path = os.path.join(TEMP_DIR_PATH, subdirectory, filename[:-5])

        remove_security_bindings_data(temp_pbix_path)

        layout_file_path = os.path.join(temp_pbix_path, "Report", "Layout")

        with open(layout_file_path, "r", encoding="utf-16-le") as layout_file:
            layout_data = layout_file.read()

        for pattern, new_value in period_replacements:
            layout_data = replace_period(layout_data, pattern, new_value)
        
        with open(layout_file_path, "w", encoding="utf-16-le") as layout_file:
            layout_file.write(layout_data)
        
        zip_pbix(subdirectory, filename)
