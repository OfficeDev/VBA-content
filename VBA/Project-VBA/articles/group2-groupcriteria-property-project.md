---
title: Group2.GroupCriteria Property (Project)
ms.prod: project-server
api_name:
- Project.Group2.GroupCriteria
ms.assetid: 0c6d6412-cd7b-7b12-1740-7cd5cd38aaf1
ms.date: 06/08/2017
---


# Group2.GroupCriteria Property (Project)

Gets or sets the  **[GroupCriteria2](groupcriteria2-object-project.md)** collection representing the fields in a group definition. Read/write **GroupCriteria2**.


## Syntax

 _expression_. **GroupCriteria**

 _expression_ An expression that returns a **Group2** object.


## Example

The following example lists all of the group criteria in the second  **Group2** object of the **TaskGroups2** collection.


```vb
Sub ListCriteria() 

 Dim criterionNum As Integer 

 Dim criteria As GroupCriteria2 

 Dim criterion As GroupCriterion2 

 

 Set criteria = ActiveProject.TaskGroups2(2).GroupCriteria 

 criterionNum = 1 

 

 For Each criterion In criteria 

 Debug.Print "Criterion " &; criterionNum &; ", Field name: " &; criterion.FieldName 

 criterionNum = criterionNum + 1 

 Next criterion 

End Sub
```


## See also


#### Concepts


[Group2 Object](group2-object-project.md)

