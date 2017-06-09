---
title: ListBox Control, List Property Example
keywords: fm20.chm5225171
f1_keywords:
- fm20.chm5225171
ms.prod: office
ms.assetid: 14396c81-9137-7352-906c-acf70e9e77b0
ms.date: 06/08/2017
---


# ListBox Control, List Property Example

The following example swaps columns of a multicolumn  **ListBox**. The sample uses the **List** property in two ways:



1. To access and exchange individual values in the  **ListBox**. In this usage, **List** has subscripts to designate the row and column of a specified value.
    
2. To initially load the  **ListBox** with values from an array. In this usage, **List** has no subscripts.
    

To use this example, copy this sample code to the Declarations portion of a form. Make sure that the form contains a  **ListBox** named ListBox1 and a **CommandButton** named CommandButton1.



```vb
Dim MyArray(6, 3) 
'Array containing column values for ListBox. 
 
Private Sub UserForm_Initialize() 
 Dim i As Single 
 
 ListBox1.ColumnCount = 3 
'This list box contains 3 data columns 
 
 'Load integer values MyArray 
 For i = 0 To 5 
 MyArray(i, 0) = i 
 MyArray(i, 1) = Rnd 
 MyArray(i, 2) = Rnd 
 Next i 
 
 'Load ListBox1 
 ListBox1.List() = MyArray 
 
End Sub
```




```vb
Private Sub CommandButton1_Click() 
' Exchange contents of columns 1 and 3 
 
 Dim i As Single 
 Dim Temp As Single 
 
 For i = 0 To 5 
 Temp = ListBox1.List(i, 0) 
 ListBox1.List(i, 0) = ListBox1.List(i, 2) 
 ListBox1.List(i, 2) = Temp 
 Next i 
End Sub
```


