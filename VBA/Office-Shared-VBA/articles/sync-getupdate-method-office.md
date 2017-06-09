---
title: Sync.GetUpdate Method (Office)
keywords: vbaof11.chm277007
f1_keywords:
- vbaof11.chm277007
ms.prod: office
api_name:
- Office.GetUpdate
ms.assetid: a92c0096-fcf2-2754-31e6-2b20a5841463
ms.date: 06/08/2017
---


# Sync.GetUpdate Method (Office)

Compares the local version of the shared document to the version on the server.


 **Note**  Beginning with Microsoft Office 2010, this object or member has been deprecated and should not be used.


## Syntax

 _expression_. **GetUpdate**

 _expression_ A variable that represents a **Sync** object.


## Remarks

Use the  **GetUpdate** method to compare the local version of the shared document to the version on the server and to refresh the sync status.

Not all document synchronization problems raise trappable run-time errors. After performing an operation using the  **Sync** object, it's a good idea to check the **Status** property; if the **Status** property is **msoSyncStatusError**, check the **ErrorType** property for additional information on the type of error that has occurred.

In many circumstances, the best way to resolve an error condition is to call the  **GetUpdate** method. For example, if a call to **PutUpdate** results in an error condition, then a call to **GetUpdate** will reset the status to **msoSyncStatusLocalChanges**.


## Example

The following example compares the local and server copies of the document using the  **GetUpdate** method and reports whether the server has a newer copy.


```
    Dim objSync As Office.Sync 
    Dim strStatus As String 
    Set objSync = ActiveDocument.Sync 
    objSync.GetUpdate 
    If objSync.Status = msoSyncStatusNewerAvailable Then 
        strStatus = "A newer version is available on the server." 
        MsgBox strStatus, vbInformation + vbOKOnly, "Sync Information" 
    End If 
    Set objSync = Nothing 

```


