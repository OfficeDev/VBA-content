
# Can't rename with different drive (Error 74)

 **Last modified:** July 28, 2015

The  **Name** statement must rename the file to the current drive. This error has the following cause and solution:




- You tried to move a file to a different drive using the  **Name** statement. Use **FileCopy** to write the file to another drive, and then delete the old file with a **Kill** statement.
    

For additional information, select the item in question and press F1 (in Windows) or HELP (on the Macintosh).
