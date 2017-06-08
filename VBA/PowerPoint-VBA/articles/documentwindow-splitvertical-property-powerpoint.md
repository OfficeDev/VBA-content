---
title: DocumentWindow.SplitVertical Property (PowerPoint)
keywords: vbapp10.chm511024
f1_keywords:
- vbapp10.chm511024
ms.prod: powerpoint
api_name:
- PowerPoint.DocumentWindow.SplitVertical
ms.assetid: 8a26332f-d00d-9816-30e1-48411db07a62
ms.date: 06/08/2017
---


# DocumentWindow.SplitVertical Property (PowerPoint)

Returns or sets the percentage of the document window height that the slide pane occupies in normal view. Corresponds to the pane divider position between the slide and notes panes. Read/write.


## Syntax

 _expression_. **SplitVertical**

 _expression_ A variable that represents a **DocumentWindow** object.


### Return Value

Long


## Remarks

The minimum value of the  **SplitVertical** property is always greater than 0% because the slide pane has a minimum height that depends on a 10% zoom level. The actual minimum value may vary depending on the size of the application window.


## Example

The following example sets the horizontal pane divider for the active document window to divide at 60% slide pane and 40% notes pane.


```vb
ActiveWindow.SplitVertical = 60
```


## See also


#### Concepts


[DocumentWindow Object](documentwindow-object-powerpoint.md)


