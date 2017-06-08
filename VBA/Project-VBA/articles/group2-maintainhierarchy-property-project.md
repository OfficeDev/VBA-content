---
title: Group2.MaintainHierarchy Property (Project)
keywords: vbapj.chm132400
f1_keywords:
- vbapj.chm132400
ms.prod: project-server
api_name:
- Project.Group2.MaintainHierarchy
ms.assetid: 47706f83-abd6-dd6b-0dff-41e260cf1107
ms.date: 06/08/2017
---


# Group2.MaintainHierarchy Property (Project)

Gets or sets a value that specifies whether hierarchy is maintained in the group view. Read/write  **Boolean**.


## Syntax

 _expression_. **MaintainHierarchy**

 _expression_ An expression that returns a **Group2** object.


## Remarks

The  **MaintainHierarchy** property corresponds to the **Maintain Hierarchy in Current Group** option in the **Group by** drop-down list on the **View** tab of the Project Ribbon.


## Example

The following example displays the name of the second  **Group2** object in the **TaskGroups2** collection, and then displays the setting for the **MaintainHierarchy** property in the **Immediate** window.


```vb
Debug.Print ActiveProject.TaskGroups2(2).Name 

Debug.Print ActiveProject.TaskGroups2(2).MaintainHierarchy
```


## See also


#### Concepts


[Group2 Object](group2-object-project.md)

