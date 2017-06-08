---
title: Project.AutoFilter Property (Project)
ms.prod: project-server
api_name:
- Project.Project.AutoFilter
ms.assetid: 3e6960f7-8a8a-6300-d74b-4e009fbcfca2
ms.date: 06/08/2017
---


# Project.AutoFilter Property (Project)

Gets or sets whether the AutoFilter feature is turned on for a project. Read/write  **Boolean**.


## Syntax

 _expression_. **AutoFilter**

 _expression_ A variable that represents a **Project** object.


## Remarks

The  **AutoFilter** setting is in the **Filter** drop-down list on the **View** tab of the Ribbon.


## Example

The following example turns on AutoFilter in the active project.


```vb
Sub turnOnAutoFilter() 
    ActiveProject.AutoFilter = True
End Sub
```


