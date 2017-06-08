---
title: Sync.LastSyncTime Property (Office)
keywords: vbaof11.chm277003
f1_keywords:
- vbaof11.chm277003
ms.prod: office
api_name:
- Office.Sync.LastSyncTime
ms.assetid: d85af059-a39e-e100-c81a-06265b43cade
ms.date: 06/08/2017
---


# Sync.LastSyncTime Property (Office)

Gets the date and time when the local copy of the active document was last synchronized with the server copy. Read-only.


 **Note**  Beginning with Microsoft Office 2010, this object or member has been deprecated and should not be used.


## Syntax

 _expression_. **LastSyncTime**

 _expression_ A variable that represents a **Sync** object.


## Remarks

Use the  **LastSyncTime** property to determine how much time has elapsed since the local copy of the active document was last synchronized with the server copy. Check the **[Status](sync-status-property-office.md)** property to determine whether the local copy and the server copy are out of sync.

If the active document is not configured for synchronization between the local copy and the server copy, the  **LastSyncTime** property raises a run-time error.


## Example

The following example alerts the user and displays the sync status if more than 24 hours have elapsed since the LastSyncTime.


```
    Dim objSync As Office.Sync 
    Dim dtmLastSync As Date 
    Dim strStatus As String 
    Set objSync = ActiveDocument.Sync 
    dtmLastSync = CDate(objSync.LastSyncTime) 
    If DateDiff("h", dtmLastSync, Now) > 24 Then 
        strStatus = "Document has not been synced " &amp; _ 
            " within the last 24 hours." &amp; vbCrLf &amp; _ 
            "Document status: " &amp; objSync.Status 
        MsgBox strStatus, vbInformation + vbOKOnly, "Error Information" 
    End If 
    Set objSync = Nothing 

```


## See also


#### Concepts


[Sync Object](sync-object-office.md)
#### Other resources


[Sync Object Members](sync-members-office.md)

