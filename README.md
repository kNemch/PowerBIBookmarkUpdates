# PBI Bookmarks Monthly Update
The script updates the Year, Month and Quarter values in PBI report's slicers and bookmarks.
### BUSINESS LOGIC of the update process:
Every 15th day from the beginning of the month the default slicers values should be changed in the reports and bookmarks.
The new default values should be the ones that represent the previous month.

### HOW TO USE THE SCRIPT
You can use the script either with or without the CLI arguments.
- Directory argument (`-d --directory`):
  * Enter the argument to specify the custom location of the "workspace" directory.
  * Ignore the argument, if the script is in the "workspace" directory.
- Year and Month arguments (`-y --year` & `-m --month`)
  * Enter the arguments to specify the custom new values of Year and Month.
    The Year and Month arguments should be passed together.
  * Ignore the arguments, to update the report according to current date.
