---
title: Sync.WorkspaceLastChangedBy Property (Office)
keywords: vbaof11.chm277002
f1_keywords:
- vbaof11.chm277002
ms.prod: office
api_name:
- Office.Sync.WorkspaceLastChangedBy
ms.assetid: f2eac8a6-5e94-44a9-3d2f-1ca04cf54361
ms.date: 06/08/2017
---


# Sync.WorkspaceLastChangedBy Property (Office)

Displays the display name of the user who last saved changes to the server copy of a shared document. Read-only.


 **Note**  Beginning with Microsoft Office 2010, this object or member has been deprecated and should not be used.


## Syntax

 _expression_. **WorkspaceLastChangedBy**

 _expression_ A variable that represents a **Sync** object.


## Remarks

If the active document is not configured for synchronization between the local copy and the server copy, the  **WorkspaceLastChangedBy** property raises a run-time error.


## Example

The following example checks for a conflict between the local and the server copies of the shared document and reports the name of the user who last saved changes to the server copy.


```
    Dim objSync As Office.Sync 
    Dim strStatus As String 
    Set objSync = ActiveDocument.Sync 
    If objSync.Status = msoSyncStatusConflict Then 
        strStatus = "The server copy has been changed." &amp; vbCrLf &amp; _ 
            "Changes have been made by: " &amp; _ 
            objSync.WorkspaceLastChangedBy 
        MsgBox strStatus, vbInformation + vbOKOnly, "Server Copy Changed" 
    End If 
    Set objSync = Nothing 

```


## See also


#### Concepts


[Sync Object](sync-object-office.md)
#### Other resources


[Sync Object Members](sync-members-office.md)

