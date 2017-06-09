---
title: Presentation.Sync Property (PowerPoint)
keywords: vbapp10.chm583084
f1_keywords:
- vbapp10.chm583084
ms.prod: powerpoint
api_name:
- PowerPoint.Presentation.Sync
ms.assetid: aebb519d-ffb8-88a8-3771-5edb6b28792c
ms.date: 06/08/2017
---


# Presentation.Sync Property (PowerPoint)

Returns a  **Sync** object that enables you to manage the synchronization of the local and server copies of a shared presentation stored in a Microsoft SharePoint Server shared workspace. Read-only.


## Syntax

 _expression_. **Sync**

 _expression_ A variable that represents a **Presentation** object.


### Return Value

Sync


## Remarks

The  **Status** property of the **Sync** object returns important information about the current state of synchronization. Use the **GetUpdate** method to refresh the sync status. Use the **LastSyncTime**, **ErrorType**, and **WorkspaceLastChangedBy** properties to return additional information.

For more information on the differences and conflicts that can exist between the local and server copies of shared presentations, see the  **Status** property.

Use the  **PutUpdate** method to save local changes to the server. Close and re-open the document to retrieve the latest version from the server when no local changes have been made. Use the **ResolveConflict** method to resolve differences between the local and the server copies, or the **OpenVersion** method to open a different version along with the currently open local version of the document.

The  **GetUpdate**, **PutUpdate**, and **ResolveConflict** methods of the **Sync** object do not return status codes because they complete their tasks asynchronously. The **Sync** object provides important status information by firing a single event, called the **PresentationSync** event of the **Application** object.

The  **PresentationSync** event returns one of the following **MsoSyncEventType** constants.


||
|:-----|
|**msoSyncEventDownloadInitiated**|
|**msoSyncEventDownloadSucceeded**|
|**msoSyncEventDownloadFailed**|
|**msoSyncEventUploadInitiated**|
|**msoSyncEventUploadSucceeded**|
|**msoSyncEventUploadFailed**|
|**msoSyncEventDownloadNoChange**|
|**msoSyncEventOffline**|
The  **Sync** object model is available whether sharing and synchronization are enabled or disabled on the active document. The **Sync** property of the **Presentation** object does not return **Nothing** when the active document is not shared or synchronization is not enabled. Use the **Status** property to determine whether the document is shared and whether synchronization is enabled.

Not all document synchronization problems raise run-time errors that can be trapped. After using the methods of the  **Sync** object, it is a good idea to check the **Status** property. If the **Status** property value is **msoSyncStatusError**, check the **ErrorType** property for additional information on the type of error that has occurred.

In many circumstances, the recommended way to resolve an error condition is to call the  **GetUpdate** method. For example, if a call to **PutUpdate** results in an error condition, a call to **GetUpdate** will reset the status to **msoSyncStatusLocalChanges**.


## Example

The following example displays the name of the last person to modify the active presentation if the active presentation is a shared document in a Document Workspace.


```vb
Dim eStatus As MsoSyncStatusType
Dim strLastUser As String

eStatus = ActivePresentation.Sync.Status

If eStatus = msoSyncStatusLatest Then
    strLastUser = ActivePresentation.Sync.WorkspaceLastChangedBy
    MsgBox "You have the most up-to-date copy." &; _
        "This file was last modified by " &; strLastUser
End If
```


 **Note**  This object or member has been deprecated, but it remains part of the object model for backward compatibility. You should not use it in new applications.


## See also


#### Concepts


[Presentation Object](presentation-object-powerpoint.md)

