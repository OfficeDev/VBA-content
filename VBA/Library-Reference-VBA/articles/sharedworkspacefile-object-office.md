---
title: SharedWorkspaceFile Object (Office)
keywords: vbaof11.chm266000
f1_keywords:
- vbaof11.chm266000
ms.prod: MULTIPLEPRODUCTS
api_name:
- Office.SharedWorkspaceFile
ms.assetid: 44e0bbfa-145d-df71-928f-2333b54f1829
---


# SharedWorkspaceFile Object (Office)

The  **SharedWorkspaceFile** object represents a file saved in a shared document workspace.


 **Note**  Beginning with Microsoft Office 2010, this object or member has been deprecated and should not be used.


## Remarks

Use the  **SharedWorkspaceFile** object to manage documents and files saved in a shared workspace.


## Example

Although the  **SharedWorkspaceFile** object has a **URL** property that returns the file's complete path and filename, it does not have a **FileName** property. Use a simple function to extract the filename from the file's URL as in the following example. An additional supporting function decodes escaped space characters in the URL.


```vb
Private Function FilenameFromURL(FileURL As String) As String 
    Dim intLastSeparator As Integer 
    FileURL = URLDecode(FileURL) 
    intLastSeparator = InStrRev(FileURL, "/") 
    FilenameFromURL = Right(FileURL, Len(FileURL) - intLastSeparator) 
End Function 
 
Private Function URLDecode(URLtoDecode As String) As String 
    URLDecode = Replace(URLtoDecode, "%20", " ") 
End Function
```

Use the  **Item** ( _index_ ) property of the **SharedWorkspaceFiles** collection to return a specific **SharedWorkspaceFile** object. Use the **CreatedBy**, **CreatedDate**, **ModifiedBy**, and **ModifiedDate** properties to return information about the history of each file. The following example returns the number of files in the shared workspace and information about each file, using the supporting functions shown above.




```vb
    Dim swsFile As Office.SharedWorkspaceFile 
    Dim strFileInfo As String 
    strFileInfo = "The shared workspace contains " &; _ 
    ActiveWorkbook.SharedWorkspace.Files.Count &; " File(s)." &; vbCrLf 
    For Each swsFile In ActiveWorkbook.SharedWorkspace.Files 
        strFileInfo = strFileInfo &; FilenameFromURL(swsFile.URL) &; vbCrLf &; _ 
            " - URL: " &; swsFile.URL &; vbCrLf &; _ 
            " - Created by: " &; swsFile.CreatedBy &; vbCrLf &; _ 
            " - Created on: " &; swsFile.CreatedDate &; vbCrLf &; _ 
            " - Modified by: " &; swsFile.ModifiedBy &; vbCrLf &; _ 
            " - Modified on: " &; swsFile.ModifiedDate &; vbCrLf 
    Next 
    MsgBox strFileInfo, vbInformation + vbOKOnly, _ 
        "Files in Shared Workspace" 
    Set swsFile = Nothing 

```


## See also


#### Concepts


[Object Model Reference](../../Office-Shared-VBA/articles/reference-object-library-reference-for-office.md)

