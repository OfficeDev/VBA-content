---
title: Application.NewPresentation Property (PowerPoint)
keywords: vbapp10.chm502049
f1_keywords:
- vbapp10.chm502049
ms.prod: powerpoint
api_name:
- PowerPoint.Application.NewPresentation
ms.assetid: 9685db30-9d73-19ad-432b-8d79b2d6ee50
ms.date: 06/08/2017
---


# Application.NewPresentation Property (PowerPoint)

Returns a  **NewFile** object that represents a presentation listed on the **New Presentation** task pane. Read-only.


## Syntax

 _expression_. **NewPresentation**

 _expression_ A variable that represents an **Application** object.


### Return Value

NewFile


## Example

This example lists a presentation on the  **New Presentation** task pane at the bottom of the last section in the pane.


```vb
Sub CreateNewPresentationListItem()

    Application.NewPresentation.Add FileName:="C:\Presentation.ppt"

    Application.CommandBars("Task Pane").Visible = True

End Sub
```


## See also


#### Concepts


[Application Object](application-object-powerpoint.md)

