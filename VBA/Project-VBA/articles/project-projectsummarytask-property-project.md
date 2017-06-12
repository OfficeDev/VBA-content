---
title: Project.ProjectSummaryTask Property (Project)
ms.prod: project-server
api_name:
- Project.Project.ProjectSummaryTask
ms.assetid: 88603abc-e988-9ab3-dc83-c44221da13b9
ms.date: 06/08/2017
---


# Project.ProjectSummaryTask Property (Project)

Gets a  **[Task](task-object-project.md)** object representing the project summary task for the active project. Read-only **Task**.


## Syntax

 _expression_. **ProjectSummaryTask**

 _expression_ A variable that represents a **Project** object.


## Remarks

Local custom fields can be accessed directly from  **ProjectSummaryTask** as task custom fields.


 **Note**  Enterprise project fields are available when a project is opened from Project Server. To access enterprise custom fields, it is necessary to use the  **SetField** and **GetField** methods along with the **[FieldNameToFieldConstant](application-fieldnametofieldconstant-method-project.md)** method.


## Example

The following example sets the local  **Cost1** task custom field and displays it in a message box.


```vb
Sub AddEnterpriseProjectCost1Value() 
    ActiveProject.ProjectSummaryTask.Cost1 = "500.00" 
 
    MsgBox "The Cost1 custom field for the project is: " _
       &; ActiveProject.ProjectSummaryTask.Cost1 
End Sub
```


