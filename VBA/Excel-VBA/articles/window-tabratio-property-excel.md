---
title: Window.TabRatio Property (Excel)
keywords: vbaxl10.chm356116
f1_keywords:
- vbaxl10.chm356116
ms.prod: excel
api_name:
- Excel.Window.TabRatio
ms.assetid: 41033d2d-9967-3990-b739-61c0649c24f3
ms.date: 06/08/2017
---


# Window.TabRatio Property (Excel)

Returns or sets the ratio of the width of the workbook's tab area to the width of the window's horizontal scroll bar (as a number between 0 (zero) and 1; the default value is 0.6). Read/write  **Double** .


## Syntax

 _expression_ . **TabRatio**

 _expression_ A variable that represents a **Window** object.


## Remarks

This property has no effect when  **[DisplayWorkbookTabs](window-displayworkbooktabs-property-excel.md)** is set to **False** (its value is retained, but it has no effect on the display).


## Example

This example makes the workbook tabs half the width of the horizontal scroll bar.


```vb
ActiveWindow.TabRatio = 0.5
```


## See also


#### Concepts


[Window Object](window-object-excel.md)

