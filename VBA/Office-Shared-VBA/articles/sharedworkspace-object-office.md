---
title: SharedWorkspace Object (Office)
keywords: vbaof11.chm276000
f1_keywords:
- vbaof11.chm276000
ms.prod: office
api_name:
- Office.SharedWorkspace
ms.assetid: 7512f0ff-382d-d344-9424-aa10549d14f9
ms.date: 06/08/2017
---


# SharedWorkspace Object (Office)

The  **SharedWorkspace** property of a **Document** object in Microsoft Word, a **Workbook** object in Microsoft Excel, and a **Presentation** object in Microsoft PowerPoint returns a **SharedWorkspace** object which allows the developer to add the active document to a SharePoint site and to manage other objects in the shared workspace site.


 **Note**  Beginning with Microsoft Office 2010, this object or member has been deprecated and should not be used.


## Remarks

Use the  **SharedWorkspace** object to add the active Word, Excel or PowerPoint document to a SharePoint document workspace site on the server to take advantage of the workspace's collaboration features, or to disconnect or remove the document from the workspace site. Use the **SharedWorkspace** object's collections to manage files, folders, links, members and tasks associated with the shared document.

The  **SharedWorkspace** object model is available whether or not a document is stored in a workspace. The **SharedWorkspace** property of the **Document**, **Workbook**, and **Presentation** objects does not return **Nothing** when the document is not shared. Use the **Connected** property of the **SharedWorkspace** object to determine whether the active document is in fact saved in and connected to a shared workspace.

Users require appropriate permissions to use the objects, properties and methods in the  **SharedWorkspace** object hierarchy.

Use the  **SharedWorkspaceFiles** collection, accessed through the **Files** property of the **SharedWorkspace** object, to manage documents and files saved in a shared workspace.

Use the  **SharedWorkspaceFolders** collection, accessed through the **Folders** property of the **SharedWorkspace** object, to manage subfolders within the main document library folder of a shared workspace.

Use the  **SharedWorkspaceLinks** collection, accessed through the **Links** property of the **SharedWorkspace** object, to manage links to additional documents and information of interest to the members who are collaborating on the documents in the shared workspace.

Use the  **SharedWorkspaceMembers** collection, accessed through the **Members** property of the **SharedWorkspace** object, to manage users who have rights to participate in a shared workspace and to collaborate on the shared documents saved in the workspace.

Use the  **SharedWorkspaceTasks** collection, accessed through the **Tasks** property of the **SharedWorkspace** object, to manage tasks assigned to the members who are collaborating on the documents in the shared workspace.

Use the  **CreateNew** method to create a new document workspace and to add the active document to the workspace. Use the **Name** and **URL** properties to return information about the workspace.

The  **SharedWorkspace** object uses a local cache of objects and properties from the server. The developer may need to update this cache before performing certain operations, or to save cached property changes back to the server. Use the **Refresh** method of the **SharedWorkspace** object to refresh the local cache from the server, and the **LastRefreshed** property to determine when the refresh operation last took place. Use the **Save** method of the **SharedWorkspaceLink** and **SharedWorkspaceTask** objects after modifying their properties locally, to upload the changes to the server.

Use the  **Disconnect** method to disconnect the local copy of the active document from the shared workspace, while leaving the shared copy intact in the workspace. Use the **RemoveDocument** method to remove the shared document from the shared workspace entirely.

Users require appropriate permissions to use the objects, properties and methods in the  **SharedWorkspace** object hierarchy. Use the **Role** argument when adding members to the **SharedWorkspaceMembers** collection to specify the set of permissions specific to each workspace member.



When using the  **SharedWorkspace** object model, it is possible to create conditions where the **SharedWorkspace** object cache is not synchronized with the user interface displayed in the **Shared Workspace** pane of the active document. For example, if the **CreateNew** method programmatically adds the active document to a new workspace while the **Shared Workspace** pane is open, the **Shared Workspace** pane will continue to display the **Create** button. In circumstances like these, if the user makes a selection in the **Shared Workspace** pane that is no longer valid, an error is raised and a refresh operation is carried out to synchronize the display with the current document state and shared workspace data.

The  **Document**, **Workbook**, and **Presentation** objects also have a **Sync** property which returns a **Sync** object. Use the **Sync** object and its properties and methods to manage the synchronization of the local and the server copies of the shared document.


## Example

The following example displays the properties of the shared workspace to which the active document is connected.


```
    Dim swsWorkspace As Office.SharedWorkspace 
    Dim strSWSInfo As String 
    Set swsWorkspace = ActiveWorkbook.SharedWorkspace 
    strSWSInfo = swsWorkspace.Name &amp; vbCrLf &amp; _ 
        " - URL: " &amp; swsWorkspace.URL &amp; vbCrLf &amp; _ 
        "The shared workspace contains " &amp; vbCrLf &amp; _ 
        " - Files: " &amp; swsWorkspace.Files.Count &amp; vbCrLf &amp; _ 
        " - Folders: " &amp; swsWorkspace.Folders.Count &amp; vbCrLf &amp; _ 
        " - Links: " &amp; swsWorkspace.Links.Count &amp; vbCrLf &amp; _ 
        " - Members: " &amp; swsWorkspace.Members.Count &amp; vbCrLf &amp; _ 
        " - Tasks: " &amp; swsWorkspace.Tasks.Count &amp; vbCrLf 
    MsgBox strSWSInfo, vbInformation + vbOKOnly, _ 
        "Shared Workspace Information" 
    Set swsWorkspace = Nothing
```


## Methods



|**Name**|
|:-----|
|[CreateNew](sharedworkspace-createnew-method-office.md)|
|[Delete](sharedworkspace-delete-method-office.md)|
|[Disconnect](sharedworkspace-disconnect-method-office.md)|
|[Refresh](sharedworkspace-refresh-method-office.md)|
|[RemoveDocument](sharedworkspace-removedocument-method-office.md)|

## Properties



|**Name**|
|:-----|
|[Application](sharedworkspace-application-property-office.md)|
|[Connected](sharedworkspace-connected-property-office.md)|
|[Creator](sharedworkspace-creator-property-office.md)|
|[Files](sharedworkspace-files-property-office.md)|
|[Folders](sharedworkspace-folders-property-office.md)|
|[LastRefreshed](sharedworkspace-lastrefreshed-property-office.md)|
|[Links](sharedworkspace-links-property-office.md)|
|[Members](sharedworkspace-members-property-office.md)|
|[Name](sharedworkspace-name-property-office.md)|
|[Parent](sharedworkspace-parent-property-office.md)|
|[SourceURL](sharedworkspace-sourceurl-property-office.md)|
|[Tasks](sharedworkspace-tasks-property-office.md)|
|[URL](sharedworkspace-url-property-office.md)|

## See also


#### Other resources


[Object Model Reference](http://msdn.microsoft.com/library/499c789a-aba2-0fad-649a-0ea964cd3b5e%28Office.15%29.aspx)
