---
title: DocumentLibraryVersion Members (Office)
ms.prod: office
ms.assetid: 81015690-f681-67e5-4ff7-329a95f78f3d
ms.date: 06/08/2017
---


# DocumentLibraryVersion Members (Office)
The  **DocumentLibraryVersion** object represents a single saved version of a shared document which has versioning enabled and which is stored in a document library on the server. Each **DocumentLibraryVersion** object is a member of the active document's **DocumentLibraryVersions** collection.

The  **DocumentLibraryVersion** object represents a single saved version of a shared document which has versioning enabled and which is stored in a document library on the server. Each **DocumentLibraryVersion** object is a member of the active document's **DocumentLibraryVersions** collection.


## Methods



|**Name**|**Description**|
|:-----|:-----|
|[Delete](documentlibraryversion-delete-method-office.md)|Removes a document library version from the  **DocumentLibraryVersions** collection.|
|[Open](documentlibraryversion-open-method-office.md)|Opens the specified version of the shared document from the  **DocumentLibraryVersions** collection in read-only mode.|
|[Restore](documentlibraryversion-restore-method-office.md)|Restores a previous saved version of a shared document from the  **DocumentLibraryVersions** collection.|

## Properties



|**Name**|**Description**|
|:-----|:-----|
|[Application](documentlibraryversion-application-property-office.md)|Gets an  **Application** object that represents the container application for the **DocumentLibraryVersion** object (you can use this property with an **Automation** object to return that object's container application). Read-only.|
|[Comments](documentlibraryversion-comments-property-office.md)|Gets any optional comments associated with the specified version of the shared document. Read-only.|
|[Creator](documentlibraryversion-creator-property-office.md)|Gets a 32-bit integer that indicates the application in which the  **DocumentLibraryVersion** object was created. Read-only.|
|[Index](documentlibraryversion-index-property-office.md)|Gets a  **Long** representing the index number for a **DocumentLibraryVersion** object in the collection. Read-only.|
|[Modified](documentlibraryversion-modified-property-office.md)|Gets the date and time at which the specified version of the shared document was last saved to the server. Read-only.|
|[ModifiedBy](documentlibraryversion-modifiedby-property-office.md)|Gets the name of the user who last saved the specified version of the shared document to the server. Read-only.|
|[Parent](documentlibraryversion-parent-property-office.md)|Gets the  **Parent** object for the **DocumentLibraryVersion** object. Read-only.|

