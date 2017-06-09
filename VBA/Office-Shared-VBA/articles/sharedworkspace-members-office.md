---
title: SharedWorkspace Members (Office)
ms.prod: office
ms.assetid: e4c2b518-d955-27e1-3e73-173d3c4f961d
ms.date: 06/08/2017
---


# SharedWorkspace Members (Office)
The  **SharedWorkspace** property of a **Document** object in Microsoft Word, a **Workbook** object in Microsoft Excel, and a **Presentation** object in Microsoft PowerPoint returns a **SharedWorkspace** object which allows the developer to add the active document to a SharePoint site and to manage other objects in the shared workspace site.

The  **SharedWorkspace** property of a **Document** object in Microsoft Word, a **Workbook** object in Microsoft Excel, and a **Presentation** object in Microsoft PowerPoint returns a **SharedWorkspace** object which allows the developer to add the active document to a SharePoint site and to manage other objects in the shared workspace site.


## Methods



|**Name**|**Description**|
|:-----|:-----|
|[CreateNew](sharedworkspace-createnew-method-office.md)|Creates a document workspace site on the server and adds the active document to the new shared workspace site.|
|[Delete](sharedworkspace-delete-method-office.md)|Deletes the current shared workspace and all data within it.|
|[Disconnect](sharedworkspace-disconnect-method-office.md)|Disconnects the local copy of the active document from the shared workspace site.|
|[Refresh](sharedworkspace-refresh-method-office.md)|Refreshes the local cache of the [SharedWorkspace](sharedworkspace-object-office.md) object's files, folders, links, members, and tasks from the server.|
|[RemoveDocument](sharedworkspace-removedocument-method-office.md)|Removes the active document from the shared workspace site.|

## Properties



|**Name**|**Description**|
|:-----|:-----|
|[Application](sharedworkspace-application-property-office.md)|Gets an  **Application** object that represents the container application for the **SharedWorkspace** object (you can use this property with an **Automation** object to return that object's container application). Read-only.|
|[Connected](sharedworkspace-connected-property-office.md)|Gets a  **Boolean** value that indicates whether or not the active document is currently saved in and connected to a shared workspace. Read-only.|
|[Creator](sharedworkspace-creator-property-office.md)|Gets a 32-bit integer that indicates the application in which the  **SharedWorkspace** object was created. Read-only.|
|[Files](sharedworkspace-files-property-office.md)|Provides access to the  **SharedWorkspaceFile** objects in the **SharedWorkspace**. Read-only.|
|[Folders](sharedworkspace-folders-property-office.md)|Gets a  **[SharedWorkspaceFolders](sharedworkspacefolders-object-office.md)** collection that represents the list of subfolders in the document library associated with the current shared workspace. Read-only.|
|[LastRefreshed](sharedworkspace-lastrefreshed-property-office.md)|Gets the date and time when the  **Refresh** method was most recently called. Read-only.|
|[Links](sharedworkspace-links-property-office.md)|Gets a  **[SharedWorkspaceLinks](sharedworkspacelinks-object-office.md)** collection that represents the list of links saved in the current shared workspace. Read-only.|
|[Members](sharedworkspace-members-property-office.md)|Gets a  **[SharedWorkspaceMembers](sharedworkspacemembers-object-office.md)** collection that represents the list of members in the current shared workspace. Read-only.|
|[Name](sharedworkspace-name-property-office.md)|Gets or sets the display name of the shared workspace site. Read/write.|
|[Parent](sharedworkspace-parent-property-office.md)|Gets the  **Parent** object for the **SharedWorkspace** object. Read-only.|
|[SourceURL](sharedworkspace-sourceurl-property-office.md)|Designates the location of the public copy of a shared document to which changes should be published back after the document has been revised in a separate document workspace site. Read-only.|
|[Tasks](sharedworkspace-tasks-property-office.md)|Gets a  **[SharedWorkspaceTasks](sharedworkspacetasks-object-office.md)** collection that represents the list of tasks in the current shared workspace. Read-only.|
|[URL](sharedworkspace-url-property-office.md)|Gets the top-level Uniform Resource Locator (URL) of the shared workspace. Read-only.|

