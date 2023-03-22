import argparse
import os
import shutil
import re
import calendar
import datetime
import zipfile

from zipfile import ZipFile


"""GLOBAL VARIABLES"""
"""Path constants"""
# The values for these global path variables are assigned in the setup_work_dir_paths() function, depending on the CLI user input
WORK_DIR_PATH:str = None    # Working directory with all the data
ORIGINALS_DIR_PATH:str = None     # Directory for original files
TEMP_DIR_PATH:str = None    # Directory for temporary (unarchived) files
#RESULTS_DIR_PATH:str = None # Directory for the modified .pbix files


"""FUNCTIONS"""
# CLI
def get_cli_parser():
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
        \r  * Enter the argument to specify the custom location of the root working directory.
        \r  * Ignore the argument, if the script is called from the working directory.
        \r- Year and Month arguments (-y --year & -m --month)
        \r  * Enter the arguments to specify the custom new values of Year and Month.
        \r    The Year and Month arguments should be passed together.
        \r  * Ignore the arguments, to update the report according to current date.
        \r- Old Year Value (-o --oldYearValue)
        \r  * Enter the value of the year, which should be replaced in slicers
        \r  * By default, the year value from previous month is used as the old value (for Feb 2023 - Jan 2023 - same years)
        ''',
        formatter_class=argparse.RawDescriptionHelpFormatter
    )
    parser.add_argument("-d", "--directory", type=str, help="The work directory - directory with .pbix files, where all updates will occur", required=False)
    parser.add_argument("-w", "--workspace", type=str, help="PBI Workspace directory", required=False)
    parser.add_argument("-r", "--report", type=str, help="Name of the report without .pbix extension", required=False)
    parser.add_argument("-y", "--year", type=int, help="New value for Year", required=False)
    parser.add_argument("-m", "--month", type=int, help="New value for Month", required=False)
    parser.add_argument("-o", "--oldYearValue", type=int, help="Old value for Year", required=False)
    return parser


def verify_cli_args(cli_parser: argparse.ArgumentParser):
    print("\nCLI - ARGUMENTS VERIFICATION")

    cli_args = cli_parser.parse_args()

    def cli_error_message():
        print("ERROR! Stopping the script execution. See the error below.\n")

    # directory
    if not cli_args.directory:
        print("""Work Directory:
            \rYou haven't provided any value to the -d --directory CLI argument.
            \rThe script will use the current working directory (the directory, from which the script is called) as the Working Directory with .pbix files.
            \r\nThe DEFAULT PATH is used (current script location):\n{}\n""".format(os.getcwd()))
    elif not os.path.exists(cli_args.directory):
        cli_error_message()
        raise cli_parser.error("""Directory Value Error: The path you provided for the Working Directory is invalid.
            \rPlease, check if there are no errors in the path.\n""")
    
    # workspace and report
    if cli_args.workspace and not cli_args.report:
        cli_error_message()
        raise cli_parser.error("""Arguments Combination Error: -r --report is required when -w --workspace is set.
            \rOnly the value for the Workspace argument was provided (-w --workspace).
            \rPlease, provide BOTH WORKSPACE and REPORT arguments to update the specific Report located in specific Workspace.\n""")
    elif cli_args.report and not cli_args.workspace:
        cli_error_message()
        raise cli_parser.error("""Arguments Combination Error: -w --workspace is required when -r --report is set.
            \rOnly the value for the Report argument was provided (-r --report).
            \rPlease, provide BOTH WORKSPACE and REPORT arguments to update the specific Report located in specific Workspace.\n""")
    
    # new year and month values
    if cli_args.month and (cli_args.month < 1 or cli_args.month > 12):
        cli_error_message()
        raise cli_parser.error("""Month Value Error: The Month value should be in range between 1 and 12.\n""")

    if cli_args.year and not cli_args.month:
        cli_error_message()
        raise cli_parser.error("""Arguments Combination Error: -m --month is required when -y --year is set.
            \rOnly the value for the Year argument was provided (-y --year).
            \rPlease, provide BOTH YEAR and MONTH arguments to update the report with your custom date values.\n""")
    elif not cli_args.year and cli_args.month:
        cli_error_message()
        raise cli_parser.error("""Arguments Combination Error: -y --year is required when -m --month is set.
            \rOnly the value for the Year argument was provided (-m --month).
            \rPlease, provide BOTH YEAR and MONTH arguments to update the report with your custom date values.\n""")
    print("OK! No issues with CLI arguments\n--------------------------------")



# WORKING DIRECTORY - PATHS, DIRECTORIES AND FILES
def setup_work_dir_paths(cli_work_dir_path:str):
    """Assigning the paths to the global variables"""
    global WORK_DIR_PATH, TEMP_DIR_PATH, RESULTS_DIR_PATH, ORIGINALS_DIR_PATH
    # If the CLI argument was not provided, we take the Current Working Directory as the root
    WORK_DIR_PATH = cli_work_dir_path if cli_work_dir_path else os.getcwd()
    TEMP_DIR_PATH = os.path.join(WORK_DIR_PATH, "#TEMP")
    ORIGINALS_DIR_PATH = os.path.join(WORK_DIR_PATH, "#ORIGINALS BACKUP {}"
                                      .format(datetime.datetime.now().strftime("%Y-%m-%d %H-%M-%S"))
                                      )  # The originals are moved from root folder to this folder
    #RESULTS_DIR_PATH = WORK_DIR_PATH  # The files with updates are saved to the root folder
    print("\nWORKING DIRECTORY PATHS:", WORK_DIR_PATH, TEMP_DIR_PATH, ORIGINALS_DIR_PATH, sep="\n")


def get_pbix_workspaces_and_filenames() -> list[tuple[str, str]]:
    # Scanning the Working Directory for .pbix files
    # We take the subfolder structures as Workspace names and .pbix filenames and save these pair into the list 
    # Is used when only the directory is passed to the script
    pbix_paths_and_files = []
    for directory_path, _, filenames in os.walk(WORK_DIR_PATH):
        for filename in filenames:
            if (filename[-5:] == ".pbix"
                    # excluding the child directories of the Working Directory, which are not the workspace directories 
                    and os.path.join(WORK_DIR_PATH, "#") not in directory_path):
                workspace_name = shorten_dir_path(directory_path)
                pbix_paths_and_files.append((workspace_name, filename))
    return pbix_paths_and_files


def create_directories_hierarchy(pbi_workspaces):
    print("\nCREATING THE DIRECTORIES HIERARCHY")
    print("Working Directory: {}".format(WORK_DIR_PATH))
    directories_to_create = [ORIGINALS_DIR_PATH, TEMP_DIR_PATH]
    for tech_dir in directories_to_create:
        if not os.path.exists(tech_dir):
            os.mkdir(tech_dir)
            print("{} [Created Directory]".format(os.path.basename(tech_dir)))
        else:
            print("{} [Directory Exists]".format(os.path.basename(tech_dir)))
        for pbi_ws_name in pbi_workspaces:
            pbi_ws_subdirectory_path = os.path.join(tech_dir, pbi_ws_name)
            if not os.path.exists(pbi_ws_subdirectory_path):
                os.mkdir(pbi_ws_subdirectory_path)
                print("* Created directory for [{}] workspace: .\\{}".format(
                    pbi_ws_name, 
                    shorten_dir_path(pbi_ws_subdirectory_path)))
            else:
                print("+ Directory exists for [{}] workspace: .\\{}".format(
                    pbi_ws_name, 
                    shorten_dir_path(pbi_ws_subdirectory_path)))


def shorten_dir_path(path):
    # Removing the Working Directory Path part from the path to improve the readability
    # Example: C:\Users\Kateryna_Nemchenko\Desktop\Bookmark Updates\#RESULTS\ws1 -> #RESULTS\ws1
    return path[len(WORK_DIR_PATH) + 1:]
    

# FILES AND DIRECTORIES MANIPULATIONS - MOVING, DELETING
def backup_original_file(src_pbix_file_path, backup_pbix_file_path):
    # Copying the original files from WORK_DIR to ORIGINALS_BACKUP_DIR
    shutil.copy(src_pbix_file_path, backup_pbix_file_path)
    print("Moved the original {} file to #ORIGINALS BACKUP folder".format(os.path.basename(src_pbix_file_path)))


def remove_temp_files():
    shutil.rmtree(TEMP_DIR_PATH)
    print("Removed the #TEMP directory")


# ARCHIVE OPERATIONS
def unzip_pbix(src_pbix_file_path, pbix_temp_files_path):
    with ZipFile(src_pbix_file_path, 'r') as source_archive:
        source_archive.extractall(path=pbix_temp_files_path)
        print('Extracting the files from .\\{} to .\\{}'
              .format(
                shorten_dir_path(src_pbix_file_path), 
                shorten_dir_path(pbix_temp_files_path)))


def zip_pbix(pbix_temp_files_path, result_pbix_file_path):
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
    
    print('Archiving the temp files to .\\{} from .\\{}'
          .format(
            shorten_dir_path(result_pbix_file_path), 
            shorten_dir_path(pbix_temp_files_path)))


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
    
    print("Removed the SecurityBindings file")


def get_patterns_and_replacements(cli_arg_year, cli_arg_month, cli_arg_old_year):
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

    old_value_year = cli_arg_old_year if cli_arg_old_year else previous_month.year

    # The single quotation marks must be incuded into the Month and Quarter strings, because it indicates that the values are of type String
    # Examples, how the data is represented in the Layout configs file:
    # year (number)    -> {\"Value\":\"2022L\"}
    # month (string)   -> {\"Value\":\"'Jan'\"}
    # quarter (string) -> {\"Value\":\"'Q1'\"}
    return [
        # pattern, new value
        ("'(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)'", "'{}'".format(new_value_month_abbr)),  # month eng
        ("'\d{1,2}月'",     "'{}月'".format(new_value_month)),          # month chn
        ("'Q [1-4]'",       "'Q {}'".format(new_value_quarter)),        # quarter eng (with space)
        ("'Q[1-4]'",    	"'Q{}'".format(new_value_quarter)),        # quarter eng (without space)
        ("'[1-4]季度'",     "'{}季度'".format(new_value_quarter)),      # quarter chn
        #(str(old_value_year),          str(new_value_year)),           # year
        ("{}L".format(old_value_year), "{}L".format(new_value_year))    # year with L
    ]


def create_value_expression(value):
    return r'\\"Value\\":\\"{}\\"'.format(value)


def replace_period(text, period_pattern, period_new_value):
    # The should be deffirent for search and replacement operations
    search_pattern = create_value_expression(period_pattern)
    new_value = create_value_expression(period_new_value)
    return re.sub(search_pattern, new_value, text)


def modify_layout_file(pbix_temp_files_path, patterns_and_new_values):
    print("Modifying the Layout file")

    layout_file_path = os.path.join(pbix_temp_files_path, "Report", "Layout")

    # It is crucial to keep the UTF-16-LE encoding for the Layout file.
    # Otherwise the PBI Desktop won't be able to read and open the .pbix file correclty
    with open(layout_file_path, "r", encoding="utf-16-le") as layout_file:
        layout_data = layout_file.read()

    for pattern, new_value in patterns_and_new_values:
        print("\t{} -> {}: {} matches".format(pattern, new_value, len(re.findall(pattern, layout_data))))
        layout_data = replace_period(layout_data, pattern, new_value)
    
    with open(layout_file_path, "w", encoding="utf-16-le") as layout_file:
        layout_file.write(layout_data)


"""OUTPUT FUNCTIONS (STDOUT)"""
def print_cli_input(cli_args):
    print("""
    \r========================
    \rMONTHLY BOOKMARKS UPDATE
    \r========================
    \r\nCLI - USER INPUT
    \rWork Directory: {}
    \rSpecific location of the Report in the Workspace (if provided):
    \r * Workspace: {}
    \r * Report: {}
    \rNew values for the reports: 
    \r * Year: {}
    \r * Month: {}
    \r * Old Year Value, that should be replaced: {}"""
          .format(cli_args.directory, cli_args.workspace, cli_args.report, 
                  cli_args.year, cli_args.month, cli_args.oldYearValue))


def print_patterns(patterns_and_new_values):
    print("\nPATTERNS")
    for period_pattern, period_new_value in patterns_and_new_values:
        print("{}  ->  {}".format(period_pattern, period_new_value))


def print_file_name(ws_subdir, pbix_filename):
    print("--------------------")
    print("WORKSPACE", ws_subdir, ": FILE", pbix_filename)    
    

"""MAIN FUNCTION"""
def main():
    # Reading user CLI input
    cli_parser = get_cli_parser()
    cli_args = cli_parser.parse_args()
    print_cli_input(cli_args)
    verify_cli_args(cli_parser)

    # Generating patterns and new values
    patterns_and_new_values = get_patterns_and_replacements(cli_args.year, cli_args.month, cli_args.oldYearValue)
    print_patterns(patterns_and_new_values)

    # Creating the path strings for main directories: Working (root), #ORIGINALS BACKUP, #TEMP
    setup_work_dir_paths(cli_args.directory)

    # Determining the list of the .pbix files and the workspace directories
    # Structure of pbix_workspaces_and_files: [("ws1", "r1"), ("ws2", "r2")] - list of tuples
    if cli_args.workspace and cli_args.report:
        # if the script is executed for specific report determined by user
        # the list will contain only the values pair passed by user
        report_filename = cli_args.report if cli_args.report[-5:] == ".pbix" else cli_args.report + ".pbix"
        pbix_workspaces_and_files = [(cli_args.workspace, report_filename)]
    else:
        # if the nothing was passed to the --workspace and --report args
        # the script searches the directory for available .pbix reports and their workspaces
        pbix_workspaces_and_files = get_pbix_workspaces_and_filenames()

    pbix_workpaces = set([pbix[0] for pbix in pbix_workspaces_and_files])
    create_directories_hierarchy(pbix_workpaces)

    # Stopping the script execution if no .pbix files were found in wroking directory
    print("\nTotal number of .pbix files found in Working Directory: {}".format(len(pbix_workspaces_and_files)))
    if len(pbix_workspaces_and_files) == 0:
        return
    
    # Processing of the .pbix files
    print("\nFILES PROCESSING")
    for ws_subdir, pbix_filename in pbix_workspaces_and_files:
        print_file_name(ws_subdir, pbix_filename)

        # Paths to src and backup files, temp directory. Source file will be replaced by the Result file in root directory
        pbix_file_path = os.path.join(WORK_DIR_PATH, ws_subdir, pbix_filename)
        pbix_backup_file_path = os.path.join(ORIGINALS_DIR_PATH, ws_subdir, pbix_filename)
        pbix_temp_files_path = os.path.join(TEMP_DIR_PATH, ws_subdir, pbix_filename[:-5])

        # Creating the file backup - moving the original file to the #ORIGINALS BACKUP
        backup_original_file(pbix_file_path, pbix_backup_file_path)
        
        # Processing of the selected file
        unzip_pbix(pbix_file_path, pbix_temp_files_path)  # unzipping from root/#ORIGINALS BACKUP to root/#TEMP
        remove_security_bindings_data(pbix_temp_files_path)  # removing check sum data
        modify_layout_file(pbix_temp_files_path, patterns_and_new_values)  # replacing the slicers values
        zip_pbix(pbix_temp_files_path, pbix_file_path)  # zipping file from root/#TEMP to root
    
    #remove_temp_files()
    #print("\n----\nDONE\n----")
    print("Please, check the {} folder for result files\n".format(WORK_DIR_PATH))


if __name__ == "__main__":
    main()
