---
title: "How to: Simulate a Two-Dimensional Array with List Boxes"
keywords: olfm10.chm3077165
f1_keywords:
- olfm10.chm3077165
ms.prod: outlook
ms.assetid: da0dd724-ff6e-04e0-c421-6011bffa750e
ms.date: 06/08/2017
---


# How to: Simulate a Two-Dimensional Array with List Boxes

The following example loads a two-dimensional array with data and, in turn, loads two  **[ListBox](listbox-object-outlook-forms-script.md)** controls using the **[Column](listbox-column-property-outlook-forms-script.md)** and **[List](listbox-list-property-outlook-forms-script.md)** properties. Note that the **Column** property transposes the array elements during loading.

To use this example, copy this sample code to the Script Editor of a form. To run the code you need to open the form so the  **Open** event will activate. Make sure that the form contains two **ListBox** controls named ListBox1 and ListBox2.



```vb
Dim MyArray(6,3) 
 
Sub Item_Open() 
 Dim i 
 
 Set ListBox1 = Item.GetInspector.ModifiedFormPages("P.2").ListBox1 
 Set ListBox2 = Item.GetInspector.ModifiedFormPages("P.2").ListBox2 
 
 ListBox1.ColumnCount = 3 'The 1st list box contains 3 data columns 
 ListBox2.ColumnCount = 6 'The 2nd box contains 6 data columns 
 
 'Load integer values into first column of MyArray 
 For i = 0 To 5 
 MyArray(i, 0) = i 
 Next 
 
 'Load columns 2 and three of MyArray 
 MyArray(0, 1) = "Zero" 
 MyArray(1, 1) = "One" 
 MyArray(2, 1) = "Two" 
 MyArray(3, 1) = "Three" 
 MyArray(4, 1) = "Four" 
 MyArray(5, 1) = "Five" 
 
 MyArray(0, 2) = "Zero" 
 MyArray(1, 2) = "Un ou Une" 
 MyArray(2, 2) = "Deux" 
 MyArray(3, 2) = "Trois" 
 MyArray(4, 2) = "Quatre" 
 MyArray(5, 2) = "Cinq" 
 
 'Load data into ListBox1 and ListBox2 
 ListBox1.List() = MyArray 
 ListBox2.Column() = MyArray 
 
End Sub
```


