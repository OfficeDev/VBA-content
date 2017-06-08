---
title: View.ReadingLayoutActualView Property (Word)
keywords: vbawd10.chm161808434
f1_keywords:
- vbawd10.chm161808434
ms.prod: word
api_name:
- Word.View.ReadingLayoutActualView
ms.assetid: 6d6b382b-21b6-79dc-31ce-6d25f70732c4
ms.date: 06/08/2017
---


# View.ReadingLayoutActualView Property (Word)

Sets or returns a  **Boolean** that represents whether pages displayed in reading layout view are displayed using the same layout as printed pages.


## Syntax

 _expression_ . **ReadingLayoutActualView**

 _expression_ An expression that returns a **[View](view-object-word.md)** object.


## Remarks

In reading layout view, pages are not displayed with the full content contained in the literal printed pages, as you would see in normal view or in print layout view. Instead they are displayed in screens. When the  **ReadingLayoutActualView** property is set to **True** , the document is displayed as it would appear when printed. On smaller monitors, this requires a zoom level that makes the document hard to read, but it is fine for larger monitors.


## Example

The following example displays the pages in reading layout view as they would appear if they were printed.


```vb
ActiveWindow.View.ReadingLayout = True 
ActiveWindow.View.ReadingLayoutActualView = True
```


## See also


#### Concepts


[View Object](view-object-word.md)

