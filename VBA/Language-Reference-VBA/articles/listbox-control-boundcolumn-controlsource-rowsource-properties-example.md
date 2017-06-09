---
title: ListBox Control, BoundColumn, ControlSource, RowSource Properties Example
keywords: fm20.chm5225190
f1_keywords:
- fm20.chm5225190
ms.prod: office
ms.assetid: 5aa015d5-0a2b-ca93-940c-3faf4dd9d900
ms.date: 06/08/2017
---


# ListBox Control, BoundColumn, ControlSource, RowSource Properties Example

The following example uses a range of worksheet cells in a  **ListBox** and, when the user selects a row from the list, displays the row index in another worksheet cell. This code sample uses the **RowSource**, **BoundColumn**, and **ControlSource** properties.

To use this example, copy this sample code to the Declarations portion of a form. Make sure that the form contains a  **ListBox** named ListBox1. In the worksheet, enter data in cells A1:E4. You also need to make sure cell A6 contains no data.



```vb
Private Sub UserForm_Initialize() 
 
ListBox1.ColumnCount = 5 
ListBox1.RowSource = "a1:e4" 
 
ListBox1.ControlSource = "a6" 
'Place the ListIndex into cell a6 
ListBox1.BoundColumn = 0 
End Sub
```


