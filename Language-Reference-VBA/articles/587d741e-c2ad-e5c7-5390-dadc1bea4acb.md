
# Can't save file to TEMP directory (Error 735)

 **Last modified:** July 28, 2015

Components often need to save temporary information to disk. This error has the following cause and solution:




- Component can't find a directory named TEMP. Create a directory named TEMP and set the TEMP environment variable equal to its path.
    
- The drive or partition containing the TEMP directory lacks sufficient space to save information. Make some space on the drive by erasing unnecessary files, or create a TEMP directory on another partition and set the TEMP environment variable equal to its path.
    

For additional information, select the item in question and press F1 (in Windows) or HELP (on the Macintosh).
