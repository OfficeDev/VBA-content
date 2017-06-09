---
title: Presentation.SharedWorkspace Property (PowerPoint)
keywords: vbapp10.chm583083
f1_keywords:
- vbapp10.chm583083
ms.prod: powerpoint
api_name:
- PowerPoint.Presentation.SharedWorkspace
ms.assetid: 79ba29b0-e51b-2644-60d7-6a044a9a7291
ms.date: 06/08/2017
---


# Presentation.SharedWorkspace Property (PowerPoint)

Returns a  **SharedWorkspace** object that represents the Document Workspace in which a specified presentation is located. Read-only.


## Syntax

 _expression_. **SharedWorkspace**

 _expression_ A variable that represents a **Presentation** object.


### Return Value

SharedWorkspace


## Remarks

Use the  **SharedWorkspace** object to add the active Microsoft PowerPoint presentation to a Microsoft SharePoint Server document workspace site on the server to take advantage of the workspace's collaboration features, or to disconnect or remove the document from the workspace. Use the **SharedWorkspace** object's collections to manage files, folders, links, members and tasks associated with the shared document.

The  **SharedWorkspace** object model is available whether or not a document is stored in a workspace. The **SharedWorkspace** property of the **Presentation** object does not return **Nothing** when the document is not shared. Use the **Connected** property of the **SharedWorkspace** object to determine whether the active presentation is in fact saved in and connected to a shared workspace.

Users require appropriate permissions to use the objects, properties, and methods in the  **SharedWorkspace** object hierarchy.

Use the  **SharedWorkspaceFiles** collection, accessed through the **Files** property of the **SharedWorkspace** object, to manage presentations and files saved in a shared workspace.

Use the  **SharedWorkspaceFolders** collection, accessed through the **Folders** property of the **SharedWorkspace** object, to manage subfolders within the main document library folder of a shared workspace.

Use the  **SharedWorkspaceLinks** collection, accessed through the **Links** property of the **SharedWorkspace** object, to manage links to additional documents and information of interest to the members who are collaborating on the documents in the shared workspace.

Use the  **SharedWorkspaceMembers** collection, accessed through the **Members** property of the **SharedWorkspace** object, to manage users who have rights to participate in a shared workspace and to collaborate on the shared documents saved in the workspace.

Use the  **SharedWorkspaceTasks** collection, accessed through the **Tasks** property of the **SharedWorkspace** object, to manage tasks assigned to the members who are collaborating on the documents in the shared workspace.

Use the  **CreateNew** method to create a new document workspace and to add the active document to the workspace. Use the **Name** and **URL** properties to return information about the workspace.

The  **SharedWorkspace** object uses a local cache of objects and properties from the server. The developer may need to update this cache before performing certain operations or to save cached property changes back to the server. Use the **Refresh** method of the **SharedWorkspace** object to refresh the local cache from the server, and the **LastRefreshed** property to determine when the refresh operation last took place. Use the **Save** method of the **SharedWorkspaceLink** and **SharedWorkspaceTask** objects after modifying their properties locally, to upload the changes to the server.

Use the  **Disconnect** method to disconnect the local copy of the active document from the shared workspace, while leaving the shared copy intact in the workspace. Use the **RemoveDocument** method to remove the shared document from the shared workspace entirely.

Users require appropriate permissions to use the objects, properties, and methods in the  **SharedWorkspace** object hierarchy. Use the Role argument when adding members to the **SharedWorkspaceMembers** collection to specify the set of permissions specific to each workspace member.

When using the  **SharedWorkspace** object model, it is possible to create conditions where the **SharedWorkspace** object cache is not synchronized with the user interface displayed in the **Shared Workspace** pane of the active document. For example, if the **CreateNew** method programmatically adds the active document to a new workspace while the **Shared Workspace** pane is open, the **Shared Workspace** pane will continue to display the **Create New** button. In circumstances like these, if the user makes a selection in the **Shared Workspace** pane that is no longer valid, an error is raised and a refresh operation is carried out to synchronize the display with the current document state and shared workspace data.

The  **Presentation** object also has a **Sync** property which returns a **Sync** object. Use the **Sync** object and its properties and methods to manage the synchronization of the local and the server copies of the shared document.


## Example

The following example returns a reference to the document workspace in which the active presentation is stored. This example assumes that the active document belongs to a document workspace.


```vb
Dim objWorkspace As SharedWorkspace

Set objWorkspace = ActivePresentation.SharedWorkspace


```


 **Note**  This object or member has been deprecated, but it remains part of the object model for backward compatibility. You should not use it in new applications.


## See also


#### Concepts


[Presentation Object](presentation-object-powerpoint.md)

