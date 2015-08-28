
# SharedWorkspace.Files Property (Office)

 **Last modified:** July 28, 2015

Provides access to the  **SharedWorkspaceFile** objects in the **SharedWorkspace**. Read-only.

 **Note**  Beginning with Microsoft Office 2010, this object or member has been deprecated and should not be used.


## Syntax

 _expression_. **Files**

 _expression_A variable that represents a  ** [SharedWorkspace](7512f0ff-382d-d344-9424-aa10549d14f9.md)** object.


## Example

The following example lists the files saved in the current shared workspace.


```
    Dim swsFiles As Office.SharedWorkspaceFiles 
    Set swsFiles = ActiveWorkbook.SharedWorkspace.Files 
    MsgBox "There are " &amp; swsFiles.Count &amp; _ 
        " file(s) 
        vbInformation + vbOKOnly, _ 
        "Collection Information" 
    Set swsFiles = Nothing 

```


## See also


#### Concepts


 [SharedWorkspace Object](7512f0ff-382d-d344-9424-aa10549d14f9.md)
#### Other resources


 [SharedWorkspace Object Members](e4c2b518-d955-27e1-3e73-173d3c4f961d.md)
