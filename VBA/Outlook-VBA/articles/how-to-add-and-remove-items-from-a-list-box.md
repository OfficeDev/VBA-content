---
title: "How to: Add and Remove Items from a List Box"
keywords: olfm10.chm3077205
f1_keywords:
- olfm10.chm3077205
ms.prod: outlook
ms.assetid: 4cff205b-4a15-d528-6ebd-adca6711a4d4
ms.date: 06/08/2017
---


# How to: Add and Remove Items from a List Box

The following example adds and deletes the contents of a  **[ListBox](listbox-object-outlook-forms-script.md)** using the **[AddItem](listbox-additem-method-outlook-forms-script.md)**,  **[RemoveItem](listbox-removeitem-method-outlook-forms-script.md)**, and  **SetFocus** methods, and the **[ListIndex](listbox-listindex-property-outlook-forms-script.md)** and **[ListCount](listbox-listcount-property-outlook-forms-script.md)** properties.


 **Note**  The  **SetFocus** method is inherited from the Microsoft Forms 2.0 **ListBox** control.


To use this example, copy this sample code to the Script Editor of a form. To run the code you need to open the form so the  **Open** event will activate. Make sure that the form contains:


- A  **ListBox** named ListBox1.
    
- Two  **[CommandButton](commandbutton-object-outlook-forms-script.md)** controls named CommandButton1 and CommandButton2.
    



```vb
Dim EntryCount 
Dim Listbox1 
 
Sub Item_Open() 
 Set ListBox1 = Item.GetInspector.ModifiedFormPages("P.2").ListBox1 
 Set CommandButton1 = Item.GetInspector.ModifiedFormPages("P.2").CommandButton1 
 Set CommandButton2 = Item.GetInspector.ModifiedFormPages("P.2").CommandButton2 
 EntryCount = 0 
 CommandButton1.Caption = "Add Item" 
 CommandButton2.Caption = "Remove Item" 
End Sub 
 
Sub CommandButton1_Click() 
 EntryCount = EntryCount + 1 
 ListBox1.AddItem (EntryCount &; " - Selection") 
End Sub 
 
 
Sub CommandButton2_Click() 
 ListBox1.SetFocus 
 
 'Ensure ListBox contains list items 
 If ListBox1.ListCount >= 1 Then 
 'If no selection, choose last list item. 
 If ListBox1.ListIndex = -1 Then 
 ListBox1.ListIndex = ListBox1.ListCount - 1 
 End If 
 ListBox1.RemoveItem (ListBox1.ListIndex) 
 End If 
End Sub
```


