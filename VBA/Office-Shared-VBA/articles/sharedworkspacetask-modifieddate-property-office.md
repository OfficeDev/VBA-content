---
title: SharedWorkspaceTask.ModifiedDate Property (Office)
keywords: vbaof11.chm264010
f1_keywords:
- vbaof11.chm264010
ms.prod: office
api_name:
- Office.SharedWorkspaceTask.ModifiedDate
ms.assetid: 26b96d4d-b3ee-a9cc-2a00-73457820b3e1
ms.date: 06/08/2017
---


# SharedWorkspaceTask.ModifiedDate Property (Office)

Gets the date and time when the  **SharedWorkspaceTask** object was last modified. Read-only.


 **Note**  Beginning with Microsoft Office 2010, this object or member has been deprecated and should not be used.


## Syntax

 _expression_. **ModifiedDate**

 _expression_ A variable that represents a **SharedWorkspaceTask** object.


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


[SharedWorkspaceTask Object](sharedworkspacetask-object-office.md)
#### Other resources


[SharedWorkspaceTask Object Members](sharedworkspacetask-members-office.md)

