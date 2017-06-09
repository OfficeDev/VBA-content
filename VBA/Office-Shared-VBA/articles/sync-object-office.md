---
title: Sync Object (Office)
keywords: vbaof11.chm277000
f1_keywords:
- vbaof11.chm277000
ms.prod: office
api_name:
- Office.Sync
ms.assetid: 1cb049a0-a803-969a-7923-15ddb8da8f3b
ms.date: 06/08/2017
---


# Sync Object (Office)

The  **Sync** property of the **Document** object in Microsoft Word, the **Workbook** object in Microsoft Excel, and the **Presentation** object in Microsoft PowerPoint returns a **Sync** object.


 **Note**  Beginning with Microsoft Office 2010, this object or member has been deprecated and should not be used.


## Remarks

Use the  **Sync** object to manage the synchronization of the local and server copies of a shared document stored in a SharePoint site. The **Status** property returns important information about the current state of synchronization. Use the **GetUpdate** method to refresh the sync status. Use the **LastSyncTime**, **ErrorType**, and **WorkspaceLastChangedBy** properties to return additional information.

See the  **Status** property for additional information on the differences and conflicts that can exist between the local and server copies of shared documents.

Use the  **PutUpdate** method to save local changes to the server. Close and re-open the document to retrieve the latest version from the server when no local changes have been made. Use the **ResolveConflict** method to resolve differences between the local and the server copies, or the **OpenVersion** method to open a different version alongside the currently open local version of the document.

The  **GetUpdate**, **PutUpdate**, and **ResolveConflict** methods of the **Sync** object do not return status codes because they complete their tasks asynchronously. The **Sync** object provides important status information through a single event, which the developer can access through the following application-specific events:


- in Word, through the  **Sync** event of the **Document** object or the **DocumentSync** event of the **Application** object;
    
- in Excel, through the  **Sync** event of the **Workbook** object or the **WorkbookSync** event of the **Application** object;
    
- in Microsoft PowerPoint, through the  **PresentationSync** event of the **Application** object.
    


The  **Sync** event described above returns an **msoSyncEventType** value.

The  **Sync** object model is available whether sharing and synchronization are enabled or disabled on the active document. The **Sync** property of the **Document**, **Workbook**, and **Presentation** objects does not return **Nothing** when the active document is not shared or synchronization is not enabled. Use the **Status** property to determine whether the document is shared and whether synchronization is enabled.

Not all document synchronization problems raise trappable run-time errors. After using the methods of the  **Sync** object, it's a good idea to check the **Status** property; if the **Status** property is **msoSyncStatusError**, check the **ErrorType** property for additional information on the type of error that has occurred.

In many circumstances, the best way to resolve an error condition is to call the  **GetUpdate** method. For example, if a call to **PutUpdate** results in an error condition, then a call to **GetUpdate** will reset the status to **msoSyncStatusLocalChanges**.


## Example

The following example demonstrates various methods of the  **Sync** object based on the status of the active document.


```
Dim objSync As Office.Sync 
    Dim strStatus As String 
    Set objSync = ActiveDocument.Sync 
    If objSync.Status > msoSyncStatusNoSharedWorkspace Then 
        Select Case objSync.Status 
            Case msoSyncStatusConflict 
                objSync.ResolveConflict msoSyncConflictMerge 
                ActiveDocument.Save 
                objSync.ResolveConflict msoSyncConflictClientWins 
                strStatus = "Conflict resolved by merging changes." 
            Case msoSyncStatusError 
                strStatus = "Last error type: " &amp; objSync.ErrorType 
            Case msoSyncStatusLatest 
                strStatus = "Document copies already in sync." 
            Case msoSyncStatusLocalChanges 
                objSync.PutUpdate 
                strStatus = "Local changes saved to server." 
            Case msoSyncStatusNewerAvailable 
                objSync.GetUpdate 
                strStatus = "Local copy updated from server." 
            Case msoSyncStatusSuspended 
                objSync.Unsuspend 
                strStatus = "Synchronization resumed." 
        End Select 
    Else 
        strStatus = "Not a shared workspace document." 
    End If 
    MsgBox strStatus, vbInformation + vbOKOnly, "Sync Information" 
    Set objSync = Nothing
```


## Methods



|**Name**|
|:-----|
|[GetUpdate](sync-getupdate-method-office.md)|
|[OpenVersion](sync-openversion-method-office.md)|
|[PutUpdate](sync-putupdate-method-office.md)|
|[ResolveConflict](sync-resolveconflict-method-office.md)|
|[Unsuspend](sync-unsuspend-method-office.md)|

## Properties



|**Name**|
|:-----|
|[Application](sync-application-property-office.md)|
|[Creator](sync-creator-property-office.md)|
|[ErrorType](sync-errortype-property-office.md)|
|[LastSyncTime](sync-lastsynctime-property-office.md)|
|[Parent](sync-parent-property-office.md)|
|[Status](sync-status-property-office.md)|
|[WorkspaceLastChangedBy](sync-workspacelastchangedby-property-office.md)|

## See also


#### Other resources


[Object Model Reference](http://msdn.microsoft.com/library/499c789a-aba2-0fad-649a-0ea964cd3b5e%28Office.15%29.aspx)
