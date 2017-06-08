---
title: "How to: Set the Item at the Top of a List and the Item that Has Focus in the List"
keywords: olfm10.chm3077260
f1_keywords:
- olfm10.chm3077260
ms.prod: outlook
ms.assetid: 3e44c81b-cf21-a1a8-dfc8-6063db095358
ms.date: 06/08/2017
---


# How to: Set the Item at the Top of a List and the Item that Has Focus in the List

The following example identifies the top item displayed in a  **[ListBox](listbox-object-outlook-forms-script.md)** and the item that has the focus within the **ListBox**. This example uses the  **[TopIndex](listbox-topindex-property-outlook-forms-script.md)** property to identify the item displayed at the top of the **ListBox** and the **[ListIndex](listbox-listindex-property-outlook-forms-script.md)** property to identify the item that has the focus. The user selects an item in the **ListBox**. The displayed values of  **TopIndex** and **ListIndex** are updated when the user selects an item or when the user clicks the **[CommandButton](commandbutton-object-outlook-forms-script.md)**.

To use this example, copy this sample code to the Script Editor of a form. To run the code you need to open the form so the  **Open** event will activate. Make sure that the form contains:

- A  **[Label](label-object-outlook-forms-script.md)** named Label1.
    
- A  **[TextBox](textbox-object-outlook-forms-script.md)** named TextBox1.
    
- A  **Label** named Label2.
    
- A  **TextBox** named TextBox2.
    
- A  **CommandButton** named CommandButton1.
    
- A  **ListBox** named ListBox1 that is bound to the Subject field.
    



```vb
Sub CommandButton1_Click() 
 Set ListBox1 = Item.GetInspector.ModifiedFormPages("P.2").Controls("ListBox1") 
 Set TextBox1 = Item.GetInspector.ModifiedFormPages("P.2").Controls("TextBox1") 
 Set TextBox2 = Item.GetInspector.ModifiedFormPages("P.2").Controls("TextBox2") 
 
 ListBox1.TopIndex = ListBox1.ListIndex 
 TextBox1.Text = ListBox1.TopIndex 
 TextBox2.Text = ListBox1.ListIndex 
End Sub 
 
Sub Item_PropertyChange(byval pname) 
 if pname = "Subject" then 
 Set TextBox1 = Item.GetInspector.ModifiedFormPages("P.2").Controls("TextBox1") 
 Set TextBox2 = Item.GetInspector.ModifiedFormPages("P.2").Controls("TextBox2") 
 Set ListBox1 = Item.GetInspector.ModifiedFormPages("P.2").Controls("ListBox1") 
 
 TextBox1.Text = ListBox1.TopIndex 
 TextBox2.Text = ListBox1.ListIndex 
 end if 
End Sub 
 
Sub Item_Open() 
 Set ListBox1 = Item.GetInspector.ModifiedFormPages("P.2").Controls("ListBox1") 
 Set CommandButton1 = Item.GetInspector.ModifiedFormPages("P.2").Controls("CommandButton1") 
 Set TextBox1 = Item.GetInspector.ModifiedFormPages("P.2").Controls("TextBox1") 
 Set TextBox2 = Item.GetInspector.ModifiedFormPages("P.2").Controls("TextBox2") 
 Set Label1 = Item.GetInspector.ModifiedFormPages("P.2").Controls("Label1") 
 Set Label2 = Item.GetInspector.ModifiedFormPages("P.2").Controls("Label2") 
 
 For i = 0 To 24 
 ListBox1.AddItem "Choice " &; (i + 1) 
 Next 
 ListBox1.Height = 66 
 CommandButton1.Caption = "Move to top of list" 
 CommandButton1.AutoSize = True 
 CommandButton1.TakeFocusOnClick = False 
 
 Label1.Caption = "Index of top item" 
 TextBox1.Text = ListBox1.TopIndex 
 
 Label2.Caption = "Index of current item" 
 Label2.AutoSize = True 
 Label2.WordWrap = False 
 TextBox2.Text = ListBox1.ListIndex 
End Sub
```


