---
title: Application.SynchronizeWithSite Method (Project)
keywords: vbapj.chm2287
f1_keywords:
- vbapj.chm2287
ms.prod: project-server
api_name:
- Project.Application.SynchronizeWithSite
ms.assetid: 1bd749d2-fe3f-ee86-dc27-5e39267901bc
ms.date: 06/08/2017
---


# Application.SynchronizeWithSite Method (Project)

Synchronizes a local project in Project Professional with a SharePoint 2013 tasks list, or synchronizes with a SharePoint task lists project that is visible in Project Web App.


## Syntax

 _expression_. **SynchronizeWithSite**( _SiteURL_,  _ListName_)

 _expression_ An expression that returns an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _SiteURL_|Optional|**String**|URL of the SharePoint site.|
| _ListName_|Optional|**String**|Name of the tasks list. For a local project, Project Professional creates the tasks list if it does not exist.|

### Return Value

 **Boolean**


## Remarks

The  **SynchronizeWithSite** method is available in Project Professional only, for a local project or for a SharePoint tasks list project that is stored in Project Web App. Saving a local project to a SharePoint site is a way to share some project details with people who do not have access to Project Web App. The SharePoint tasks list also enables users who have the correct permission to add tasks, assign tasks to resources, set task priority (low, normal, or high), set task status and percent complete, and set task precedessors.

For a tasks list project that SharePoint manages, when you use Project Professional to open that project from Project Web App, you can synchronize changes with the SharePoint tasks list manually in the Backstage view, or programmatically by using the  **SynchronizeWithSite** method.

If a resource assigned to a task does not exist in the SharePoint farm, or if more than one resource is assigned to a task, the resources cannot be published to the tasks list. However, the resources remain assigned to tasks in the project plan. Project shows another dialog box that explains the resource issue. When the user chooses  **OK**, Project creates the specified tasks list.


 **Tip**  To create a local project that uses resources available in a SharePoint site, it is easiest to create the project without local resources, use the SharePoint tasks list to add resources, and then use Project to synchronize with the SharePoint changes.

When changes are made to the SharePoint tasks list, running  **SynchronizeWithSite** again displays the **Conflict Resolution** dialog box, which enables you to choose the SharePoint version or the Project version of each modified task. You can also choose **Keep the selected version for all remaining conflicts in this synchronization**.

The  **SynchronizeWithSite** method corresponds to **Sync with a SharePoint Tasks List** on the **Share** tab of the Backstage view.


## Example

The following example creates a SharePoint tasks list named "Test Tasks List" on the site http://OurTeam.


```vb
Sub CreateSharePointTasksList() 
    Application.SynchronizeWithSite SiteURL:="http://OurTeam", _
        ListName:="Test Tasks List" 
End Sub
```

After you create a tasks list, it is not necessary to specifiy the SiteURL or ListName arguments again to synchronize the project with the same tasks list.




```vb
Sub SyncWithExistingTasksList() 
    Application.SynchronizeWithSite 
End Sub
```

For an example that synchronizes the  **Priority** column in a SharePoint tasks list with the **Priority** field in Project tasks, see the **[ManageSiteColumns](application-managesitecolumns-method-project.md)** method.


