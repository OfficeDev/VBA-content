---
title: TextColumn Property Example
keywords: fm20.chm5225159
f1_keywords:
- fm20.chm5225159
ms.prod: office
ms.assetid: a794e071-456b-1b5d-d02a-5130cdacb79a
ms.date: 06/08/2017
---


# TextColumn Property Example

The following example uses the  **TextColumn** property to identify the column of data in a **ListBox** that supplies data for its **Text** property. This example sets the third column of the **ListBox** as the text column. As you select an entry from the **ListBox**, the value from the **TextColumn** will be displayed in the **Label**.

This example also demonstrates how to load a multicolumn  **ListBox** using the **AddItem** method and the **List** property.

To use this example, copy this sample code to the Declarations portion of a form. Make sure that the form contains:



- A  **ListBox** named ListBox1.
    
- A  **TextBox** named TextBox1.
    




```vb
Private Sub UserForm_Initialize() 
ListBox1.ColumnCount = 3 
 
ListBox1.AddItem "Row 1, Col 1" 
ListBox1.List(0, 1) = "Row 1, Col 2" 
ListBox1.List(0, 2) = "Row 1, Col 3" 
 
ListBox1.AddItem "Row 2, Col 1" 
ListBox1.List(1, 1) = "Row 2, Col 2" 
ListBox1.List(1, 2) = "Row 2, Col 3" 
 
ListBox1.AddItem "Row 3, Col 1" 
ListBox1.List(2, 1) = "Row 3, Col 2" 
ListBox1.List(2, 2) = "Row 3, Col 3" 
 
ListBox1.TextColumn = 3 
End Sub 
 
Private Sub ListBox1_Change() 
TextBox1.Text = ListBox1.Text 
End Sub
```


