---
title: "How to: Swap Columns of a Multicolumn List Box"
keywords: olfm10.chm3077206
f1_keywords:
- olfm10.chm3077206
ms.prod: outlook
ms.assetid: 5d6fb3f2-161e-eeb6-1d0c-dc4d4670214b
ms.date: 06/08/2017
---


# How to: Swap Columns of a Multicolumn List Box

The following example swaps columns of a multicolumn  **[ListBox](listbox-object-outlook-forms-script.md)**. The sample uses the  **[List](listbox-list-property-outlook-forms-script.md)** property in two ways:


1. To access and exchange individual values in the  **ListBox**. In this usage,  **List** has subscripts to designate the row and column of a specified value.
    
2. To initially load the  **ListBox** with values from an array. In this usage, **List** has no subscripts.
    

To use this example, copy this sample code to the Script Editor of a form. To run the code you need to open the form so the  **Open** event will activate. Make sure that the form contains a **ListBox** named ListBox1 and a **[CommandButton](commandbutton-object-outlook-forms-script.md)** named CommandButton1.




```vb
Dim Listbox1 
Dim MyArray(6, 3) 
 
Sub Item_Open 
 Dim i 
 Set Listbox1 = Item.GetInspector.ModifiedFormPages("P.2").Listbox1 
 
 Listbox1.ColumnCount = 3 
 For i = 0 to 5 
 MyArray(i, 0) = i 
 MyArray(i, 1) = Rnd 
 MyArray(i, 2) = Rnd 
 Next 
 
 Listbox1.List() = MyArray 
End Sub 
 
Sub CommandButton1_Click 
 Dim i 
 Dim Temp 
 
 For i = 0 to 5 
 Temp = Listbox1.List(i, 0) 
 Listbox1.List(i, 0) = Listbox1.List(i, 2) 
 Listbox1.List(i, 2) = Temp 
 Next 
End Sub
```


