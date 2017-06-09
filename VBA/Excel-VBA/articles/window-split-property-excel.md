---
title: Window.Split Property (Excel)
keywords: vbaxl10.chm356111
f1_keywords:
- vbaxl10.chm356111
ms.prod: excel
api_name:
- Excel.Window.Split
ms.assetid: 7fcc304f-8a42-f997-2c32-5a9793683bd5
ms.date: 06/08/2017
---


# Window.Split Property (Excel)

 **True** if the window is split. Read/write **Boolean** .


## Syntax

 _expression_ . **Split**

 _expression_ A variable that represents a **Window** object.


## Remarks

It's possible for  **[FreezePanes](window-freezepanes-property-excel.md)** to be **True** and **Split** to be **False** , or vice versa.

This property applies only to worksheets and macro sheets.


## Example

This example splits the active window in Book1.xls at cell B2, without freezing panes. This causes the  **Split** property to return **True** .


```vb
Workbooks("BOOK1.XLS").Worksheets("Sheet1").Activate 
With ActiveWindow 
 .SplitColumn = 2 
 .SplitRow = 2 
End With
```

This example illustrates two ways of removing the split added by the preceding example.




```vb
Workbooks("BOOK1.XLS").Worksheets("Sheet1").Activate 
ActiveWindow.Split = False 'method one 
Workbooks("BOOK1.XLS").Worksheets("Sheet1").Activate 
ActiveWindow.SplitColumn = 0 'method two 
ActiveWindow.SplitRow = 0
```

This example removes the window split. Before you can remove the split, you must set  **FreezePanes** to **False** to remove frozen panes.




```vb
Workbooks("BOOK1.XLS").Worksheets("Sheet1").Activate 
With ActiveWindow 
 .FreezePanes = False 
 .Split = False 
End With
```


## See also


#### Concepts


[Window Object](window-object-excel.md)

