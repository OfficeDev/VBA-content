---
title: ListBox Control, AddItem, RemoveItem Methods, ListIndex, ListCount Properties Example
keywords: fm20.chm5225178
f1_keywords:
- fm20.chm5225178
ms.prod: office
ms.assetid: 70bc2f0c-79a5-89f2-e987-84f673d4bf97
ms.date: 06/08/2017
---


# ListBox Control, AddItem, RemoveItem Methods, ListIndex, ListCount Properties Example

The following example adds and deletes the contents of a  **ListBox** using the **AddItem** and **RemoveItem** methods, and the **ListIndex** and **ListCount** properties.

To use this example, copy this sample code to the Declarations portion of a form. Make sure that the form contains:




- A  **ListBox** named ListBox1.
    
- Two  **CommandButton** controls named CommandButton1 and CommandButton2.
    




```vb
Dim EntryCount As Single 
 
Private Sub CommandButton1_Click() 
 EntryCount = EntryCount + 1 
 ListBox1.AddItem (EntryCount &; " - Selection") 
End Sub
```




```vb
Private Sub CommandButton2_Click() 
 'Ensure ListBox contains list items 
 If ListBox1.ListCount >= 1 Then 
 'If no selection, choose last list item. 
 If ListBox1.ListIndex = -1 Then 
 ListBox1.ListIndex = _ 
 ListBox1.ListCount - 1 
 End If 
 ListBox1.RemoveItem (ListBox1.ListIndex) 
 End If 
End Sub
```




```vb
Private Sub UserForm_Initialize() 
 EntryCount = 0 
 CommandButton1.Caption = "Add Item" 
 CommandButton2.Caption = "Remove Item" 
End Sub
```


