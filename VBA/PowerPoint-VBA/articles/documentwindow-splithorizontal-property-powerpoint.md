---
title: DocumentWindow.SplitHorizontal Property (PowerPoint)
keywords: vbapp10.chm511025
f1_keywords:
- vbapp10.chm511025
ms.prod: powerpoint
api_name:
- PowerPoint.DocumentWindow.SplitHorizontal
ms.assetid: 89ec538b-d8a3-23e8-a246-35c44884a432
ms.date: 06/08/2017
---


# DocumentWindow.SplitHorizontal Property (PowerPoint)

Returns or sets the percentage of the document window width that the outline pane occupies in normal view. Corresponds to the pane divider position between the slide and outline panes. Read/write.


## Syntax

 _expression_. **SplitHorizontal**

 _expression_ A variable that represents a **DocumentWindow** object.


### Return Value

Long


## Remarks

The maximum value of the  **SplitHorizontal** property is always less than 100% because the slide pane has a minimum width that depends on a 10% zoom level. The actual maximum value may vary depending on the size of the application window.


## Example

The following example sets the vertical pane divider for the active document window to divide at 15% outline pane and 85% slide pane.


```vb
ActiveWindow.SplitHorizontal = 15
```


## See also


#### Concepts


[DocumentWindow Object](documentwindow-object-powerpoint.md)


