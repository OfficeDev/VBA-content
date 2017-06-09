---
title: DocumentWindow.Selection Property (PowerPoint)
keywords: vbapp10.chm511003
f1_keywords:
- vbapp10.chm511003
ms.prod: powerpoint
api_name:
- PowerPoint.DocumentWindow.Selection
ms.assetid: 0cd670b2-53a5-87d7-8b38-761920dd9758
ms.date: 06/08/2017
---


# DocumentWindow.Selection Property (PowerPoint)

Returns a  **[Selection](selection-object-powerpoint.md)** object that represents the selection in the specified document window. Read-only.


## Syntax

 _expression_. **Selection**

 _expression_ A variable that represents a **DocumentWindow** object.


### Return Value

Selection


## Example

If there's text selected in the active window, this example makes the text italic.


```vb
With Application.ActiveWindow.Selection

    If .Type = ppSelectionText Then

        .TextRange.Font.Italic = True

    End If

End With


```


## See also


#### Concepts



[DocumentWindow Object](documentwindow-object-powerpoint.md)

