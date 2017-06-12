---
title: Sync Members (Office)
ms.prod: office
ms.assetid: 748726bd-83de-425a-5af8-177c34e3a013
ms.date: 06/08/2017
---


# Sync Members (Office)
The  **Sync** property of the **Document** object in Microsoft Word, the **Workbook** object in Microsoft Excel, and the **Presentation** object in Microsoft PowerPoint returns a **Sync** object.

The  **Sync** property of the **Document** object in Microsoft Word, the **Workbook** object in Microsoft Excel, and the **Presentation** object in Microsoft PowerPoint returns a **Sync** object.


## Methods



|**Name**|**Description**|
|:-----|:-----|
|[GetUpdate](sync-getupdate-method-office.md)|Compares the local version of the shared document to the version on the server.|
|[OpenVersion](sync-openversion-method-office.md)|Opens a different version of the shared document alongside the currently open local version.|
|[PutUpdate](sync-putupdate-method-office.md)|Updates the server copy of the shared document with the local copy.|
|[ResolveConflict](sync-resolveconflict-method-office.md)|Resolves conflicts between the local and the server copies of a shared document.|
|[Unsuspend](sync-unsuspend-method-office.md)|Resumes synchronization between the local copy and the server copy of a shared document.|

## Properties



|**Name**|**Description**|
|:-----|:-----|
|[Application](sync-application-property-office.md)|Gets an  **Application** object that represents the container application for the **Sync** object (you can use this property with an **Automation** object to return that object's container application). Read-only.|
|[Creator](sync-creator-property-office.md)|Gets a 32-bit integer that indicates the application in which the  **Sync** object was created. Read-only.|
|[ErrorType](sync-errortype-property-office.md)|Gets a  **MsoSyncErrorType** constant which indicates the type of the most recent document synchronization error. Read-only.|
|[LastSyncTime](sync-lastsynctime-property-office.md)|Gets the date and time when the local copy of the active document was last synchronized with the server copy. Read-only.|
|[Parent](sync-parent-property-office.md)|Gets the  **Parent** object for the **Sync** object. Read-only.|
|[Status](sync-status-property-office.md)|Gets the status of the synchronization of the local copy of the active document with the server copy. Read-only.|
|[WorkspaceLastChangedBy](sync-workspacelastchangedby-property-office.md)|Displays the display name of the user who last saved changes to the server copy of a shared document. Read-only.|

