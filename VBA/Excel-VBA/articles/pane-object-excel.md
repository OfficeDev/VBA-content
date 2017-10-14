---
title: Pane Object (Excel)
keywords: vbaxl10.chm359072
f1_keywords:
- vbaxl10.chm359072
ms.prod: excel
api_name:
- Excel.Pane
ms.assetid: 9064bb89-d08c-bbd3-3c0f-77a39586bbbb
ms.date: 06/08/2017
---


# Pane Object (Excel)

Represents a pane of a window.


## Remarks

 **Pane** objects exist only for worksheets and Microsoft Excel 4.0 macro sheets. The **Pane** object is a member of the **[Panes](panes-object-excel.md)** collection. The **Panes** collection contains all of the panes shown in a single window.


## Example

Use  **[Panes](window-panes-property-excel.md)** ( _index_ ), where _index_ is the pane index number, to return a single **Pane** object. The following example splits the window in which worksheet one is displayed and then scrolls through the pane in the lower-left corner until row five is at the top of the pane.


```vb
Worksheets(1).Activate 
ActiveWindow.Split = True 
ActiveWindow.Panes(3).ScrollRow = 5
```


## See also


#### Other resources


[Excel Object Model Reference](http://msdn.microsoft.com/library/11ea8598-8a20-92d5-f98b-0da04263bf2c%28Office.15%29.aspx)


