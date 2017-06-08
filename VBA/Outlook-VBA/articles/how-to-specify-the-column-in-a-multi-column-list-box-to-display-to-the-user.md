---
title: "How to: Specify the Column in a Multi-Column List Box to Display to the User"
keywords: olfm10.chm3077255
f1_keywords:
- olfm10.chm3077255
ms.prod: outlook
ms.assetid: f56b48b4-8ea7-8b77-99a1-0b522f0c9db3
ms.date: 06/08/2017
---


# How to: Specify the Column in a Multi-Column List Box to Display to the User

The following example uses the  **[TextColumn](listbox-textcolumn-property-outlook-forms-script.md)** property to identify the column of data in a **[ListBox](listbox-object-outlook-forms-script.md)** that supplies data for its **[Text](listbox-text-property-outlook-forms-script.md)** property. This example sets the third column of the **ListBox** as the text column. As you select an entry from the **ListBox**, the value from the  **TextColumn** will be displayed in the **[TextBox](textbox-object-outlook-forms-script.md)**.

This example also demonstrates how to load a multicolumn  **ListBox** using the **[AddItem](listbox-additem-method-outlook-forms-script.md)** method and the **[List](listbox-list-property-outlook-forms-script.md)** property.

To use this example, copy this sample code to the Script Editor of a form. To run the code you need to open the form so the  **Open** event will activate. Make sure that the form contains:


- A  **ListBox** named ListBox1.
    
- A  **TextBox** named TextBox1.
    



```vb
Dim ListBox1 
Dim TextBox1 
 
Sub Item_Open() 
 Set ListBox1 = Item.GetInspector.ModifiedFormPages.Item("P.2").Controls("ListBox1") 
 Set TextBox1 = Item.GetInspector.ModifiedFormPages.Item("P.2").Controls("TextBox1") 
 
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
 
Sub ListBox1_Click() 
 TextBox1.Text = ListBox1.Text 
End Sub
```


