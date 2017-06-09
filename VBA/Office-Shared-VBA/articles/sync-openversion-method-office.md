---
title: Sync.OpenVersion Method (Office)
keywords: vbaof11.chm277009
f1_keywords:
- vbaof11.chm277009
ms.prod: office
api_name:
- Office.Sync.OpenVersion
ms.assetid: 22892531-5e6d-f977-c430-0160cadb4490
ms.date: 06/08/2017
---


# Sync.OpenVersion Method (Office)

Opens a different version of the shared document alongside the currently open local version.


 **Note**  Beginning with Microsoft Office 2010, this object or member has been deprecated and should not be used.


## Syntax

 _expression_. **OpenVersion**( **_SyncVersionType_** )

 _expression_ A variable that represents a **Sync** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _SyncVersionType_|Required|**MsoSyncVersionType**| Represents the type of version.|

## Remarks

Use the  **OpenVersion** method to open the last version viewed ( **msoSyncVersionLastViewed** ) or the server copy ( **msoSyncVersionServer** ) of the shared document alongside the currently open local version.

 ** The msoSyncVersionLastViewed** option displays the copy of the document that is created whenever the user overwrites the local copy with the server copy. For example, if you call the **ResolveConflict** method with the **msoSyncConflictServerWins** option, then your local changes are saved and can be viewed by calling ** OpenVersion(msoSyncVersionLastViewed)**.

Not all document synchronization problems raise trappable run-time errors. After performing an operation using the  **Sync** object, it's a good idea to check the **Status** property; if the **Status** property is **msoSyncStatusError**, check the **ErrorType** property for additional information on the type of error that has occurred.


## Example

The following example prompts the user to open the server copy of the shared document alongside the currently open local version.


```
    Dim objSync As Office.Sync 
    Dim lngChoice As VbMsgBoxResult 
    Set objSync = ActiveDocument.Sync 
    lngChoice = MsgBox("View server copy?", _ 
        vbQuestion + vbOKCancel, "Open Server Version?") 
    If lngChoice = vbOK Then 
        objSync.OpenVersion msoSyncVersionServer 
    End If 
    Set objSync = Nothing 

```


## See also


#### Concepts


[Sync Object](sync-object-office.md)
#### Other resources


[Sync Object Members](sync-members-office.md)

