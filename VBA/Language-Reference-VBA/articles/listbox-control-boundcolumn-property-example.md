---
title: ListBox Control, BoundColumn Property Example
keywords: fm20.chm5225170
f1_keywords:
- fm20.chm5225170
ms.prod: office
ms.assetid: 17b19e02-5db4-459b-533e-73220730de01
ms.date: 06/08/2017
---


# ListBox Control, BoundColumn Property Example

The following example demonstrates how the  **BoundColumn** property influences the value of a **ListBox**. The user can choose to set the value of the **ListBox** to the index value of the specified row, or to a specified column of data in the **ListBox**.

To use this example, copy this sample code to the Declarations portion of a form. Make sure that the form contains:




- A  **ListBox** named ListBox1.
    
- A  **Label** named Label1.
    
- Three  **OptionButton** controls named OptionButton1, OptionButton2, and OptionButton3.
    




```vb
Private Sub UserForm_Initialize() 
 ListBox1.ColumnCount = 2 
 
 ListBox1.AddItem "Item 1, Column 1" 
 ListBox1.List(0, 1) = "Item 1, Column 2" 
 ListBox1.AddItem "Item 2, Column 1" 
 ListBox1.List(1, 1) = "Item 2, Column 2" 
 
 ListBox1.Value = "Item 1, Column 1" 
 
 OptionButton1.Caption = "List Index" 
 OptionButton2.Caption = "Column 1" 
 OptionButton3.Caption = "Column 2" 
 OptionButton2.Value = True 
End Sub 
 
Private Sub OptionButton1_Click() 
 ListBox1.BoundColumn = 0 
 Label1.Caption = ListBox1.Value 
End Sub 
 
Private Sub OptionButton2_Click() 
 ListBox1.BoundColumn = 1 
 Label1.Caption = ListBox1.Value 
End Sub 
 
Private Sub OptionButton3_Click() 
 ListBox1.BoundColumn = 2 
 Label1.Caption = ListBox1.Value 
End Sub 
 
Private Sub ListBox1_Click() 
 Label1.Caption = ListBox1.Value 
End Sub
```


