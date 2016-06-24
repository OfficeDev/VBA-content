
# SharedWorkspaceFolder.FolderName Property (Office)

Gets the name of a subfolder within the main document library folder of a shared workspace. Read-only.


 **Note**  Beginning with Microsoft Office 2010, this object or member has been deprecated and should not be used.


## Syntax

 _expression_. **FolderName**

 _expression_ A variable that represents a **SharedWorkspaceFolder** object.


## Remarks

The  **FolderName** property returns the subfolder name in the format parentfolder/subfolder. For example, if the shared workspace contains a folder named "Supporting Documents", the **FolderName** property returns Shared Documents/Supporting Documents.


## Example

The following example displays the number of subfolders in the shared workspace and their names.


```vb
    Dim swsFolder As Office.SharedWorkspaceFolder 
    Dim strFolderInfo As String 
    strFolderInfo = "The shared workspace contains " &; _ 
        ActiveWorkbook.SharedWorkspace.Folders.Count &; " folder(s)." &; vbCrLf 
    If ActiveWorkbook.SharedWorkspace.Folders.Count > 0 Then 
        For Each swsFolder In ActiveWorkbook.SharedWorkspace.Folders 
            strFolderInfo = strFolderInfo &; swsFolder.FolderName &; vbCrLf 
        Next 
    End If 
    MsgBox strFolderInfo, vbInformation + vbOKOnly, _ 
        "Folders in Shared Workspace" 
    Set swsFolder = Nothing 

```


## See also


#### Concepts


[SharedWorkspaceFolder Object](297c4ed7-2232-5240-ca34-d374038c66a2.md)
#### Other resources


[SharedWorkspaceFolder Object Members](e7e0a32a-ce01-e08f-f251-27d93273110e.md)