---
title: Sync.PutUpdate Method (Office)
keywords: vbaof11.chm277008
f1_keywords:
- vbaof11.chm277008
ms.prod: office
api_name:
- Office.Sync.PutUpdate
ms.assetid: 2197cb71-e4d3-e89f-768b-7fd76f92a2d2
ms.date: 06/08/2017
---


# Sync.PutUpdate Method (Office)

Updates the server copy of the shared document with the local copy.


 **Note**  Beginning with Microsoft Office 2010, this object or member has been deprecated and should not be used.


## Syntax

 _expression_. **PutUpdate**

 _expression_ A variable that represents a **Sync** object.


## Remarks

The  **PutUpdate** method can encounter a conflict condition if the client is unaware of recent changes to the server copy of the shared document. Call the **GetUpdate** method before calling **PutUpdate** to refresh the status of the server copy and to detect a possible conflict.

The  **PutUpdate** method raises a run-time error if the local document has unsaved changes.

Not all document synchronization problems raise trappable run-time errors. After performing an operation using the  **Sync** object, it's a good idea to check the **Status** property; if the **Status** property is **msoSyncStatusError**, check the **ErrorType** property for additional information on the type of error that has occurred.

In many circumstances, the best way to resolve an error condition is to call the  **GetUpdate** method. For example, if a call to **PutUpdate** results in an error condition, then a call to **GetUpdate** will reset the status to **msoSyncStatusLocalChanges**.


## Example

The following example updates the server copy of the document from the local copy using the PutUpdate method if the local copy has been edited.


```
    Dim objSync As Office.Sync 
    Dim strStatus As String 
    Set objSync = ActiveDocument.Sync 
    If objSync.Status = msoSyncStatusLocalChanges Then 
        objSync.PutUpdate 
        strStatus = "Local changes saved to server." 
        MsgBox strStatus, vbInformation + vbOKOnly, "Sync Information" 
    End If 
    Set objSync = Nothing 

```


## See also


#### Concepts


[Sync Object](sync-object-office.md)
#### Other resources


[Sync Object Members](sync-members-office.md)

