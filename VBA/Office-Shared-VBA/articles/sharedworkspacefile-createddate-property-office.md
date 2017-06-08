---
title: SharedWorkspaceFile.CreatedDate Property (Office)
keywords: vbaof11.chm266003
f1_keywords:
- vbaof11.chm266003
ms.prod: office
api_name:
- Office.SharedWorkspaceFile.CreatedDate
ms.assetid: c3a45dbd-c6b2-3046-2388-ed23ca7e36f0
ms.date: 06/08/2017
---


# SharedWorkspaceFile.CreatedDate Property (Office)

Gets the date and time when the shared workspace object was created. Read-only.


## Syntax

 _expression_. **CreatedDate**

 _expression_ A variable that represents a **SharedWorkspaceFile** object.


### Return Value

Variant


## Example

The following example returns a list of shared workspace files whose date and time created is earlier than today.


```
 Dim swsFile As Office.SharedWorkspaceFile 
 Dim dtmMidnight As Date 
 Dim dtmFileDate As Date 
 Dim strOlderFiles As String 
 dtmMidnight = CDate(FormatDateTime(Now, vbShortDate) &amp; " 12:00:00 am") 
 For Each swsFile In ActiveWorkbook.SharedWorkspace.Files 
 dtmFileDate = swsFile.CreatedDate 
 If dtmFileDate < dtmMidnight Then 
 strOlderFiles = strOlderFiles &amp; swsFile.URL &amp; vbCrLf 
 End If 
 Next 
 MsgBox "Files older than today: " &amp; vbCrLf &amp; strOlderFiles, _ 
 vbInformation + vbOKOnly, "Older Files" 
 Set swsFile = Nothing 
 

```


 **Note**  Beginning with Microsoft Office 2010, this object or member has been deprecated and should not be used.


## See also


#### Concepts


[SharedWorkspaceFile Object](sharedworkspacefile-object-office.md)
#### Other resources


[SharedWorkspaceFile Object Members](sharedworkspacefile-members-office.md)

