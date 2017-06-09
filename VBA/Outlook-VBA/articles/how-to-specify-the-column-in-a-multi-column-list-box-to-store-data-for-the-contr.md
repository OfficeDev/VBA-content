---
title: "How to: Specify the Column in a Multi-Column List Box to Store Data for the Control"
keywords: olfm10.chm3077160
f1_keywords:
- olfm10.chm3077160
ms.prod: outlook
ms.assetid: 11481d20-6c2c-2dfb-4afe-fdc4a4e1563c
ms.date: 06/08/2017
---


# How to: Specify the Column in a Multi-Column List Box to Store Data for the Control

The following example demonstrates how the  **[BoundColumn](listbox-boundcolumn-property-outlook-forms-script.md)** property influences the value of a **[ListBox](listbox-object-outlook-forms-script.md)**. The user can choose to set the value of the  **ListBox** to the index value of the specified row, or to a specified column of data in the **ListBox**.

To use this example, copy this sample code to the Script Editor of a form. To run the code you need to open the form so the  **Open** event will activate. Make sure that the form contains:

- A  **ListBox** named ListBox1.
    
- A  **[Label](label-object-outlook-forms-script.md)** named Label1.
    
- Three  **[OptionButton](optionbutton-object-outlook-forms-script.md)** controls named OptionButton1, OptionButton2, and OptionButton3.
    



```vb
Dim Listbox1 
Dim OptionButton1 
Dim OptionButton2 
Dim OptionButton3 
Dim Label1 
 
Sub Item_Open 
 Set Listbox1 = Item.GetInspector.ModifiedFormPages("P.2").Listbox1 
 Set OptionButton1 = Item.GetInspector.ModifiedFormPages("P.2").OptionButton1 
 Set OptionButton2 = Item.GetInspector.ModifiedFormPages("P.2").OptionButton2 
 Set OptionButton3 = Item.GetInspector.ModifiedFormPages("P.2").OptionButton3 
 Set Label1 = Item.GetInspector.ModifiedFormPages("P.2").Label1 
 
 Listbox1.ColumnCount = 2 
 Listbox1.AddItem "Item 1, Column 1" 
 Listbox1.List(0, 1) = "Item 1, Column 2" 
 Listbox1.AddItem "Item 2, Column 1" 
 Listbox1.List(1, 1) = "Item 2, Column 2" 
 Listbox1.Value = "Item 1, Column 1" 
 OptionButton1.Caption = "List Index" 
 OptionButton2.Caption = "Column 1" 
 OptionButton3.Caption = "Column 2" 
 OptionButton2.Value = True 
End Sub 
 
Sub OptionButton1_Click 
 Listbox1.BoundColumn = 0 
 Label1.Caption = Listbox1.Value 
End Sub 
 
Sub OptionButton2_Click 
 Listbox1.BoundColumn = 1 
 Label1.Caption = Listbox1.Value 
End Sub 
 
Sub OptionButton3_Click 
 Listbox1.BoundColumn = 2 
 Label1.Caption = Listbox1.Value 
End Sub 
 
Sub Listbox1_Click 
 Label1.Caption = Listbox1.Value 
End Sub
```


