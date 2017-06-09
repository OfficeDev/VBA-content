---
title: DocumentWindow.ActivePane Property (PowerPoint)
keywords: vbapp10.chm511022
f1_keywords:
- vbapp10.chm511022
ms.prod: powerpoint
api_name:
- PowerPoint.DocumentWindow.ActivePane
ms.assetid: 8fa4c8a1-37b6-2676-1cfd-5fa2b130d2e3
ms.date: 06/08/2017
---


# DocumentWindow.ActivePane Property (PowerPoint)

Returns a  **[Pane](pane-object-powerpoint.md)** object that represents the active pane in the document window. Read-only.


## Syntax

 _expression_. **ActivePane**

 _expression_ A variable that represents an **DocumentWindow** object.


### Return Value

Pane


## Example

If the active pane is the slide pane, this example makes the notes pane the active pane. The notes pane is the third member of the  **Panes** collection.


```vb
With ActiveWindow

    If .ActivePane.ViewType = ppViewSlide Then

        .Panes(3).Activate

    End If

End With
```


## See also


#### Concepts



[DocumentWindow Object](documentwindow-object-powerpoint.md)

