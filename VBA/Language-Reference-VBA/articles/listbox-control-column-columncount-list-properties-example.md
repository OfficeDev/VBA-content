---
title: ListBox Control, Column, ColumnCount, List Properties Example
keywords: fm20.chm5225184
f1_keywords:
- fm20.chm5225184
ms.prod: office
ms.assetid: 30706933-a979-6392-848f-1527e3ec1847
ms.date: 06/08/2017
---


# ListBox Control, Column, ColumnCount, List Properties Example

The following example loads a two-dimensional array with data and, in turn, loads two  **ListBox** controls using the **Column** and **List** properties. Note that the **Column** property transposes the array elements during loading.

To use this example, copy this sample code to the Declarations portion of a form. Make sure that the form contains two  **ListBox** controls named ListBox1 and ListBox2.



```vb
Dim MyArray(6,3) 
 
Private Sub UserForm_Initialize() 
 Dim i As Single 
 'The 1st list box contains 3 data columns 
 ListBox1.ColumnCount = 3 
 'The 2nd box contains 6 data columns 
 ListBox2.ColumnCount = 6 
 
 'Load integer values into first column of MyArray 
 For i = 0 To 5 
 MyArray(i, 0) = i 
 Next i 
 
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


