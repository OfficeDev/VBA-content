---
title: Sync.Unsuspend Method (Office)
keywords: vbaof11.chm277011
f1_keywords:
- vbaof11.chm277011
ms.prod: office
api_name:
- Office.Sync.Unsuspend
ms.assetid: 456a5f22-30bf-224d-7e3c-092711188f80
ms.date: 06/08/2017
---


# Sync.Unsuspend Method (Office)

Resumes synchronization between the local copy and the server copy of a shared document.


 **Note**  Beginning with Microsoft Office 2010, this object or member has been deprecated and should not be used.


## Syntax

 _expression_. **Unsuspend**

 _expression_ A variable that represents a **Sync** object.


## Remarks

Use the  **Unsuspend** method to resume document synchronization when the **Status** property returns **msoSyncStatusSuspended**.

Not all document synchronization problems raise trappable run-time errors. After performing an operation using the  **Sync** object, it's a good idea to check the **Status** property; if the **Status** property is **msoSyncStatusError**, check the **ErrorType** property for additional information on the type of error that has occurred.


## Example

The following example resumes document synchronization if it has been suspended.


```
    Dim objSync As Office.Sync 
    Set objSync = ActiveDocument.Sync 
    If objSync.Status = msoSyncStatusSuspended Then 
        objSync.Unsuspend 
        MsgBox "Synchronization resumed.", _ 
            vbInformation + vbOKOnly, "Sync Status" 
    End If 
    Set objSync = Nothing 

```


## See also


#### Concepts


[Sync Object](sync-object-office.md)
#### Other resources


[Sync Object Members](sync-members-office.md)

