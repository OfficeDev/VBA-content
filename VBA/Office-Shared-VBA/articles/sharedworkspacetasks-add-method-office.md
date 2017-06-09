---
title: SharedWorkspaceTasks.Add Method (Office)
keywords: vbaof11.chm265003
f1_keywords:
- vbaof11.chm265003
ms.prod: office
api_name:
- Office.SharedWorkspaceTasks.Add
ms.assetid: f427945e-e699-9ba0-6d83-98f9b77b4500
ms.date: 06/08/2017
---


# SharedWorkspaceTasks.Add Method (Office)

Adds a task to the list of tasks in a shared workspace. Returns a  **SharedWorkspaceTask** object.


 **Note**  Beginning with Microsoft Office 2010, this object or member has been deprecated and should not be used.


## Syntax

 _expression_. **Add**( **_Title_**, **_Status_**, **_Priority_**, **_Assignee_**, **_Description_**, **_Due Date_** )

 _expression_ Required. A variable that represents a **[SharedWorkspaceTasks](sharedworkspacetasks-object-office.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Title_|Required|**String**|The title of the new task.|
| _Status_|Optional|**msoSharedWorkspaceTask**|The status of the new task. Default is  **msoSharedWorkspaceTaskNotStarted**.|
| _Priority_|Optional|**msoSharedWorkspaceTask**|The priority of the new task. Default is  **msoSharedWorkspaceTaskNormal**.|
| _Assignee_|Optional|**SharedWorkspaceMember**|The member to whom the new task is assigned.|
| _Description_|Optional|**String**|The description of the new task.|
| _DueDate_|Optional|**Date**|The due date of the new task.|

## Remarks

The schema that defines shared workspace tasks and their properties for a SharePoint site can be modified on the server in such a way that the  **Add** method of the **SharedWorkspaceTasks** collection may raise an error, or may disregard the values of certain arguments. In particular, the task status and priority enumerations can be customized. Some examples of the problems that can result are mentioned below:


- If a  _Status_ argument is supplied, and the status field has been removed from the customized tasks schema, the argument will be ignored and no error will be raised.
    
- If a  _Status_ value is supplied that lies outside the status values recognized by the customized tasks schema, the argument will be ignored, the default value will be used, and no error will be raised.
    
- If a new required field has been added to the customized tasks schema, then the  **Add** method will fail with an error, and it will no longer be possible to use the **Add** method to add new tasks.
    



## Example

The following example adds a new task to the tasks collection of the shared workspace, specifies a due date, and assigns the task to the first member of the shared workspace.


```
   Dim swsTask As Office.SharedWorkspaceTask 
    Dim swsMember As Office.SharedWorkspaceMember 
    Set swsMember = ActiveWorkbook.SharedWorkspace.Members(1) 
    Set swsTask = ActiveWorkbook.SharedWorkspace.Tasks.Add( _ 
        "Complete document by year-end", _ 
        msoSharedWorkspaceTaskStatusNotStarted, _ 
        msoSharedWorkspaceTaskPriorityNormal, _ 
        swsMember, _ 
        "My first shared workspace task", #12/31/2005#) 
    MsgBox "New task: " &amp; swsTask.Title, _ 
        vbInformation + vbOKOnly, _ 
        "New Task in Shared Workspace" 
    Set swsMember = Nothing 
    Set swsTask = Nothing 

```


## See also


#### Concepts


[SharedWorkspaceTasks Object](sharedworkspacetasks-object-office.md)
#### Other resources


[SharedWorkspaceTasks Object Members](sharedworkspacetasks-members-office.md)

