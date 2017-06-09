---
title: SharedWorkspace.Refresh Method (Office)
keywords: vbaof11.chm276007
f1_keywords:
- vbaof11.chm276007
ms.prod: office
api_name:
- Office.SharedWorkspace.Refresh
ms.assetid: 62059fb9-b695-78e7-ad44-c3b918c9d423
ms.date: 06/08/2017
---


# SharedWorkspace.Refresh Method (Office)

Refreshes the local cache of the [SharedWorkspace](sharedworkspace-object-office.md) object's files, folders, links, members, and tasks from the server.


 **Note**  Beginning with Microsoft Office 2010, this object or member has been deprecated and should not be used.


## Syntax

 _expression_. **Refresh**

 _expression_ A variable that represents a **SharedWorkspace** object.


## Remarks

Use the  **Refresh** method to ensure that you are working with the most up-to-date list of objects and their properties from the shared workspace.


## Example

The following example refreshes the shared workspace if it has not been refreshed in the last 3 minutes. The example also handles the error condition where the workspace has not yet been refreshed.


```
    On Error GoTo err_NeverRefreshed 
    If DateDiff("s", ActiveWorkbook.SharedWorkspace.LastRefreshed, Now) > 180 Then 
        ActiveWorkbook.SharedWorkspace.Refresh 
    End If 
    Exit Sub 
     
err_NeverRefreshed: 
          ActiveWorkbook.SharedWorkspace.Refresh 

```


## See also


#### Concepts


[SharedWorkspace Object](sharedworkspace-object-office.md)
#### Other resources


[SharedWorkspace Object Members](sharedworkspace-members-office.md)

