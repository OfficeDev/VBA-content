---
title: Assignments.Add Method (Project)
ms.prod: project-server
api_name:
- Project.Assignments.Add
ms.assetid: c135a80e-1fb9-32e3-864e-f701c1947ca4
ms.date: 06/08/2017
---


# Assignments.Add Method (Project)

Adds an  **Assignment** object to an **Assignments** collection.


## Syntax

 _expression_. **Add**( ** _TaskID_**, ** _ResourceID_**, ** _Units_** )

 _expression_ A variable that represents an **Assignments** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _TaskID_|Optional|**Long**|The identification number of a task. Required if the parent object is a resource. The task is assigned the resource specified by ResourceID. The default value of  **TaskID** is the identification number of the parent object of the **Assignments** collection if the parent object is a **Task** object.|
| _ResourceID_|Optional|**Long**|The identification number of a resource. Required if the parent object is a task. The resource is assigned the task specified with the TaskID argument. The default value of ResourceID is the identification number of the parent object of the  **Assignments** collection if the parent object is a **Resource** object.|
| _Units_|Optional|**Variant**|The number of resource units, expressed as a decimal or percentage, to assign to the task. The default value is 1 or 100%, depending on whether the  **Show assignment units as a** setting is **Decimal** or **Percentage**, on the  **Schedule** tab of the **Project Options** dialog box. If the maximum number of units is less than 1 (or the maximum percentage is less than 100%), the default value of the Units argument is the maximum number of units (or the maximum percentage).|

### Return Value

 **Assignment**


## See also


#### Concepts


[Assignments Collection Object](assignments-object-project.md)
