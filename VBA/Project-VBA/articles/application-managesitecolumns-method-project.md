---
title: Application.ManageSiteColumns Method (Project)
keywords: vbapj.chm2288
f1_keywords:
- vbapj.chm2288
ms.prod: project-server
api_name:
- Project.Application.ManageSiteColumns
ms.assetid: 1900552c-6320-2ff5-4a07-bc6ebee60696
ms.date: 06/08/2017
---


# Application.ManageSiteColumns Method (Project)

Displays the  **Manage Fields** dialog box, which enables synchronizing built-in fields and custom fields in a local project with specified columns in a SharePoint 2013 tasks list.


## Syntax

 _expression_. **ManageSiteColumns**

 _expression_ An expression that returns an **Application** object.


### Return Value

 **Boolean**


## Remarks

The  **ManageSiteColumns** method is available only in Project Professional, with a local project that has been saved to a SharePoint task list. For more information, see the **[SynchronizeWithSite](application-synchronizewithsite-method-project.md)** method.

The following table shows the columns and default synchronized fields in the  **Manage Fields** dialog box. By default, the **Priority** and **Task Status** SharePoint columns are not synchronized with any Project field, so those items are empty.


||||
|:-----|:-----|:-----|
|**Sync**|**Project Field**|**SharePoint Column**|
|Yes|Name|Title|
|Yes|Start|Start Date|
|Yes|Finish|Due Date|
|Yes|% Complete|% Complete|
|Yes|Resource Names|Assigned To|
|Yes|Predecessors|Predecessors|
|No||Priority|
|No||Task Status|

## Example

To add the  **Priority** field in the Project Field column and synchronize with the **Priority** column in SharePoint, for example, you could do the following:


1. Rename a text custom field in Project; for example, name  **Text1** as **SharePoint Priority**.
    
2. Run the  **ManageSiteColumns** method, and then in the **Manage Fields** dialog box, select **SharePoint Priority (Text1)** in the **Project Field** drop-down list that corresponds to **Priority** in the SharePoint column.
    
3. Run the  **SyncPriority** macro.
    





```vb
Sub SyncPriority() 
    Dim tsk As Task 
    Dim msfPriority As String 
 
    Application.SynchronizeWithSite 
 
    For Each tsk In ActiveProject.Tasks 
        msfPriority = tsk.Text1 
 
        Select Case msfPriority 
            Case "(1) High" 
               tsk.Priority = 700 
           Case "(2) Normal" 
               tsk.Priority = 500 
           Case "(3) Low" 
               tsk.Priority = 300 
        End Select 
    Next tsk 
End Sub
```


