---
title: Application.DisplayScrollBars Property (Project)
ms.prod: project-server
api_name:
- Project.Application.DisplayScrollBars
ms.assetid: 4c8e2aa3-3d85-94c8-d1ce-67586b78e7e7
ms.date: 06/08/2017
---


# Application.DisplayScrollBars Property (Project)

 **True** if the scroll bars are visible for all projects. Read/write **Boolean**.


## Syntax

 _expression_. **DisplayScrollBars**

 _expression_ A variable that represents an **Application** object.


## Example

The following example changes the setting of the  **DisplayScrollBars** property.


```vb
Sub ChangeDisplayScrollBars 
 DisplayScrollBars = Not DisplayScrollBars 
End Sub
```


