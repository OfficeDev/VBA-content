---
title: MsoSyncStatusType Enumeration (Office)
ms.prod: office
api_name:
- Office.MsoSyncStatusType
ms.assetid: 52dab603-eb05-709a-99d5-908f2713b953
ms.date: 06/08/2017
---


# MsoSyncStatusType Enumeration (Office)

Specifies the status of the synchronization of the local copy of the active document with the server copy. Used with the  **Status** property of the **Sync** object.

Beginning with Microsoft Office 2010, this object or member has been deprecated and should not be used.


|**Name**|**Value**|**Description**|
|:-----|:-----|:-----|
|**msoSyncStatusConflict**|4|Both the local and the server copies have changes.|
|**msoSyncStatusError**|6|An error occurred. Use  **ErrorType** property of **Sync** object to determine exact error.|
|**msoSyncStatusLatest**|1|Documents are already in sync.|
|**msoSyncStatusLocalChanges**|3|Only local copy has changes.|
|**msoSyncStatusNewerAvailable**|2|Only server copy has changes.|
|**msoSyncStatusNoSharedWorkspace**|0|No shared workspace.|
|**msoSyncStatusNotRoaming**|0|No syncronization is needed.|
|**msoSyncStatusSuspended**|5|Syncronization has been suspended.|

