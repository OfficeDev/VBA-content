---
title: Panes Object (Excel)
keywords: vbaxl10.chm357072
f1_keywords:
- vbaxl10.chm357072
ms.prod: excel
api_name:
- Excel.Panes
ms.assetid: ce27ae27-52d9-9e51-a068-b9c082a0a692
ms.date: 06/08/2017
---


# Panes Object (Excel)

A collection of all the  **[Pane](pane-object-excel.md)** objects shown in the specified window.


## Remarks

 **Pane** objects exist only for worksheets and Microsoft Excel 4.0 macro sheets.


## Example

Use the  **Panes** property to return the **Panes** collection. The following example freezes panes in the active window if the window contains more than one pane.


```vb
If ActiveWindow.Panes.Count > 1 Then _ 
 ActiveWindow.FreezePanes = True
```

Use  **[Panes](window-panes-property-excel.md)** ( _index_ ), where _index_ is the pane index number, to return a single **Pane** object. The following example scrolls through the upper-left pane of the window in which Sheet1 is displayed.




```vb
Worksheets("sheet1").Activate 
Windows(1).Panes(1).LargeScroll down:=1
```


## See also


#### Other resources


[Excel Object Model Reference](http://msdn.microsoft.com/library/11ea8598-8a20-92d5-f98b-0da04263bf2c%28Office.15%29.aspx)


