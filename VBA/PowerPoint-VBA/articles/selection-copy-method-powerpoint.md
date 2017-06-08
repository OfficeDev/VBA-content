---
title: Selection.Copy Method (PowerPoint)
keywords: vbapp10.chm508004
f1_keywords:
- vbapp10.chm508004
ms.prod: powerpoint
api_name:
- PowerPoint.Selection.Copy
ms.assetid: 954106da-a2a9-0c55-114a-5a79f578e0c4
ms.date: 06/08/2017
---


# Selection.Copy Method (PowerPoint)

Copies the specified object to the Clipboard.


## Syntax

 _expression_. **Copy**

 _expression_ A variable that represents a **Selection** object.


## Remarks

Use the  **Paste** method to paste the contents of the Clipboard.


## Example

This example copies the selection in window one to the Clipboard and then pastes it into the view in window two. If the Clipboard contents cannot be pasted into the view in window two — for example, if you try to paste a shape into slide sorter view — this example fails.


```
Windows(1).Selection.Copy

Windows(2).View.Paste
```


## See also


#### Concepts


[Selection Object](selection-object-powerpoint.md)

