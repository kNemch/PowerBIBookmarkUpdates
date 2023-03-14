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
TEMP_DIR_PATH:str = None    # directory for temporary (unarchived) files
RESULTS_DIR_PATH:str = None # directory for the modified .pbix files


"""FUNCTIONS"""
# CLI
def get_cli_args() -> argparse.ArgumentParser:
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
    global WORK_DIR_PATH, TEMP_DIR_PATH, RESULTS_DIR_PATH
    WORK_DIR_PATH = cli_work_dir_path if cli_work_dir_path else os.getcwd()
    TEMP_DIR_PATH = os.path.join(WORK_DIR_PATH, "#TEMP")
    RESULTS_DIR_PATH = os.path.join(WORK_DIR_PATH, "#RESULTS")


def get_pbix_workspaces_and_filenames() -> list[tuple[str, str]]:
    pbix_paths_and_files = []
    for directory_path, _, filenames in os.walk(WORK_DIR_PATH):
        for filename in filenames:
            if (filename[-5:] == ".pbix" 
                    and RESULTS_DIR_PATH not in directory_path 
                    and TEMP_DIR_PATH not in directory_path):
                workspace_name = directory_path[len(WORK_DIR_PATH) + 1:]
                pbix_paths_and_files.append((workspace_name, filename))
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
def unzip_pbix(ws_subdir, pbix_filename):
    src_pbix_file_path = os.path.join(WORK_DIR_PATH, ws_subdir, pbix_filename)
    pbix_temp_files_path = os.path.join(TEMP_DIR_PATH, ws_subdir, pbix_filename[:-5])

    with ZipFile(src_pbix_file_path, 'r') as source_archive:
        source_archive.extractall(path=pbix_temp_files_path)
        print('Extracting the files from', src_pbix_file_path, "to", pbix_temp_files_path)


def zip_pbix(ws_subdir, pbix_filename):
    result_pbix_file_path = os.path.join(RESULTS_DIR_PATH, ws_subdir, pbix_filename)
    pbix_temp_files_path = os.path.join(TEMP_DIR_PATH, ws_subdir, pbix_filename[:-5])

    with ZipFile(result_pbix_file_path, mode="w") as result_archive:
        for directory_path, _, filenames in os.walk(pbix_temp_files_path):
            for filename in filenames:
                # where the file is located in the file system
                file_location_in_fs = os.path.join(directory_path, filename)
                # path for the archive hierarchy
                file_location_in_archive = file_location_in_fs[len(pbix_temp_files_path)+1:]
                # selecting the compression mode depending on archive file type
                file_compression_mode = zipfile.ZIP_STORED if filename == "DataModel" else zipfile.ZIP_DEFLATED
                result_archive.write(
                    file_location_in_fs, 
                    arcname=file_location_in_archive, 
                    compress_type=file_compression_mode)
    
    print('Writing the temporary files to', result_pbix_file_path, "from", pbix_temp_files_path)


# PBIX MODIFICATION
def remove_security_bindings_data(pbix_temp_files_path):
    """Deleting the PBI report's Control Sum Data. The control sum is located in the SecurityBindings file.
    The deletion of this data enables us to open the PBI report even if the changes were made outside Power BI Desktop"""
    
    # Removing the SecurityBindings file
    security_bindings_file_path = os.path.join(pbix_temp_files_path, "SecurityBindings")
    if os.path.exists(security_bindings_file_path):
        os.remove(security_bindings_file_path)

    # Removing the XML record about the SecurityBingings file from the [Content_Types].xml file
    # Optional step - in most cases, it is enough to delete only the SecurityBindings file
    content_types_file_path = os.path.join(pbix_temp_files_path, "[Content_Types].xml")
    with open(content_types_file_path, "r") as content_types_file:
        xml = content_types_file.read()
    updated_xml = xml.replace('<Override PartName="/SecurityBindings" ContentType="" />', "")
    with open(content_types_file_path, "w") as content_types_file: # overwriting the file
        content_types_file.write(updated_xml)


def get_patterns_and_replacements(cli_arg_year, cli_arg_month):
    # If the user didn't provide the new Year and Month values in the CLI
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

    return [
        # pattern, new value
        ('(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)', new_value_month_abbr),  # month eng
        ('\d{1,2}月', str(new_value_month) + "月"),    # month chn
        ('Q [1-4]', 'Q ' + str(new_value_quarter)),    # quarter eng (with space)
        ('Q[1-4]', 'Q' + str(new_value_quarter)),      # quarter eng (without space)
        ('[1-4]季度', str(new_value_quarter) + "季度"), # quarter chn
        (str(previous_month.year) + "L", str(new_value_year) + "L"),  # year with L
        (str(previous_month.year), str(new_value_year))  # year
    ]


def replace_period(text:str, period_pattern:str, period_new_value:str):
    search_pattern = r'\\"Value\\":\\"\'{}\'\\"'.format(period_pattern)
    new_value = r'\\"Value\\":\\"' + "'{}'".format(period_new_value) + '\\"'
    return re.sub(search_pattern, new_value, text)


def modify_layout_file(pbix_temp_files_path):
    pass



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
    for period_pattern, period_new_value in patterns_and_new_values:
        print("{}  ->  {}".format(period_pattern, period_new_value))


def print_file_name(ws_subdir, pbix_filename):
    print("--------------------")
    print("WORKSPACE", ws_subdir, ": FILE", pbix_filename)


"""MAIN FUNCTION"""
if __name__ == "__main__":

    # Reading user CLI input
    cli_args = get_cli_args()
    print_cli_input(cli_args)

    # Generating patterns and new values
    patterns_and_new_values = get_patterns_and_replacements(cli_args.year, cli_args.month)
    print_patterns(patterns_and_new_values)

    # Creating directories, getting the list of PBIX files with paths to them
    setup_work_dir_paths(cli_args.directory)
    pbix_workspaces_and_files = get_pbix_workspaces_and_filenames()
    pbix_workpaces = set([pbix[0] for pbix in pbix_workspaces_and_files])
    create_directories_hierarchy(pbix_workpaces)

    # Processing of the .pbix files
    print("\nFILES PROCESSING")
    for ws_subdir, pbix_filename in pbix_workspaces_and_files:
        print_file_name(ws_subdir, pbix_filename)

        # Paths to src and result files, temp directory
        src_pbix_file_path = os.path.join(WORK_DIR_PATH, ws_subdir, pbix_filename)
        pbix_temp_files_path = os.path.join(TEMP_DIR_PATH, ws_subdir, pbix_filename[:-5])
        result_pbix_file_path = os.path.join(RESULTS_DIR_PATH, ws_subdir, pbix_filename)

        unzip_pbix(ws_subdir, pbix_filename)

        # Working with the unpacked archive
        remove_security_bindings_data(pbix_temp_files_path)

        layout_file_path = os.path.join(pbix_temp_files_path, "Report", "Layout")

        # encoding
        with open(layout_file_path, "r", encoding="utf-16-le") as layout_file:
            layout_data = layout_file.read()

        for pattern, new_value in patterns_and_new_values:
            layout_data = replace_period(layout_data, pattern, new_value)
        
        with open(layout_file_path, "w", encoding="utf-16-le") as layout_file:
            layout_file.write(layout_data)
        
        zip_pbix(ws_subdir, pbix_filename)
    print("\n----\nDONE\n----")
