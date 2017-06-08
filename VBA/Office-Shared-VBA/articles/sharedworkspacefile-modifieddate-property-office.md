---
title: SharedWorkspaceFile.ModifiedDate Property (Office)
keywords: vbaof11.chm266005
f1_keywords:
- vbaof11.chm266005
ms.prod: office
api_name:
- Office.SharedWorkspaceFile.ModifiedDate
ms.assetid: c4d0f54c-db16-8a1e-a5d0-56ec9d5287fa
ms.date: 06/08/2017
---


# SharedWorkspaceFile.ModifiedDate Property (Office)

Gets the date and time when the  **SharedWorkspaceFile** object was last modified. Read-only.


 **Note**  Beginning with Microsoft Office 2010, this object or member has been deprecated and should not be used.


## Syntax

 _expression_. **ModifiedDate**

 _expression_ A variable that represents a **SharedWorkspaceFile** object.


### Return Value

Variant


## Example

The following example returns a list of shared workspace files whose date and time last modified is earlier than today.


```
Dim swsFile As Office.SharedWorkspaceFile 
    Dim dtmMidnight As Date 
    Dim dtmFileDate As Date 
    Dim strOlderFiles As String 
    dtmMidnight = CDate(FormatDateTime(Now, vbShortDate) &amp; " 12:00:00 am") 
    For Each swsFile In ActiveWorkbook.SharedWorkspace.Files 
        dtmFileDate = swsFile.ModifiedDate 
        If dtmFileDate < dtmMidnight Then 
            strOlderFiles = strOlderFiles &amp; swsFile.URL &amp; vbCrLf 
        End If 
    Next 
    MsgBox "Files not modified today: " &amp; vbCrLf &amp; strOlderFiles, _ 
        vbInformation + vbOKOnly, "Older Files" 
    Set swsFile = Nothing
```


## See also


#### Concepts


[SharedWorkspaceFile Object](sharedworkspacefile-object-office.md)
#### Other resources


[SharedWorkspaceFile Object Members](sharedworkspacefile-members-office.md)

