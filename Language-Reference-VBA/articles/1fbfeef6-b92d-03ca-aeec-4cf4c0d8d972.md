
# User-defined type may not be passed ByVal

 **Last modified:** July 28, 2015

 [User-defined types](b8bdf64f-5920-1ae9-16d0-b26d09524a30.md) can only be passed [by reference](b8bdf64f-5920-1ae9-16d0-b26d09524a30.md) (the default), not [by value](b8bdf64f-5920-1ae9-16d0-b26d09524a30.md). The error may not be reported until the call is made. This error has the following cause and solution:




- You placed a  **ByVal** keyword in the definition of a [parameter](b8bdf64f-5920-1ae9-16d0-b26d09524a30.md) that represented a user-defined type.
    
    Remove the  **ByVal** keyword. To keep changes from being propagated back to the caller, **Dim** a temporary [variable](b8bdf64f-5920-1ae9-16d0-b26d09524a30.md) of the type and pass the temporary variable into the [procedure](b8bdf64f-5920-1ae9-16d0-b26d09524a30.md).
    

For additional information, select the item in question and press F1 (in Windows) or HELP (on the Macintosh).
