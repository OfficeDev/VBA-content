---
title: SharedWorkspaceFile Object (Office)
keywords: vbaof11.chm266000
f1_keywords:
- vbaof11.chm266000
ms.prod: office
api_name:
- Office.SharedWorkspaceFile
ms.assetid: 44e0bbfa-145d-df71-928f-2333b54f1829
ms.date: 06/08/2017
---


# SharedWorkspaceFile Object (Office)

The  **SharedWorkspaceFile** object represents a file saved in a shared document workspace.


 **Note**  Beginning with Microsoft Office 2010, this object or member has been deprecated and should not be used.


## Remarks

Use the  **SharedWorkspaceFile** object to manage documents and files saved in a shared workspace.


## Example

Although the  **SharedWorkspaceFile** object has a **URL** property that returns the file's complete path and filename, it does not have a **FileName** property. Use a simple function to extract the filename from the file's URL as in the following example. An additional supporting function decodes escaped space characters in the URL.


```
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




```
    Dim swsFile As Office.SharedWorkspaceFile 
    Dim strFileInfo As String 
    strFileInfo = "The shared workspace contains " &amp; _ 
    ActiveWorkbook.SharedWorkspace.Files.Count &amp; " File(s)." &amp; vbCrLf 
    For Each swsFile In ActiveWorkbook.SharedWorkspace.Files 
        strFileInfo = strFileInfo &amp; FilenameFromURL(swsFile.URL) &amp; vbCrLf &amp; _ 
            " - URL: " &amp; swsFile.URL &amp; vbCrLf &amp; _ 
            " - Created by: " &amp; swsFile.CreatedBy &amp; vbCrLf &amp; _ 
            " - Created on: " &amp; swsFile.CreatedDate &amp; vbCrLf &amp; _ 
            " - Modified by: " &amp; swsFile.ModifiedBy &amp; vbCrLf &amp; _ 
            " - Modified on: " &amp; swsFile.ModifiedDate &amp; vbCrLf 
    Next 
    MsgBox strFileInfo, vbInformation + vbOKOnly, _ 
        "Files in Shared Workspace" 
    Set swsFile = Nothing 

```


## Methods



|**Name**|
|:-----|
|[Delete](sharedworkspacefile-delete-method-office.md)|

## Properties



|**Name**|
|:-----|
|[Application](sharedworkspacefile-application-property-office.md)|
|[CreatedBy](sharedworkspacefile-createdby-property-office.md)|
|[CreatedDate](sharedworkspacefile-createddate-property-office.md)|
|[Creator](sharedworkspacefile-creator-property-office.md)|
|[ModifiedBy](sharedworkspacefile-modifiedby-property-office.md)|
|[ModifiedDate](sharedworkspacefile-modifieddate-property-office.md)|
|[Parent](sharedworkspacefile-parent-property-office.md)|
|[URL](sharedworkspacefile-url-property-office.md)|

## See also


#### Other resources


[Object Model Reference](http://msdn.microsoft.com/library/499c789a-aba2-0fad-649a-0ea964cd3b5e%28Office.15%29.aspx)
