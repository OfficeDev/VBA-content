---
title: Copy, GetFromClipboard, GetText Methods, DataObject Object Example
keywords: fm20.chm5225163
f1_keywords:
- fm20.chm5225163
ms.prod: office
ms.assetid: 6be27a7e-58f2-5cad-5ed0-570520fd61f1
ms.date: 06/08/2017
---


# Copy, GetFromClipboard, GetText Methods, DataObject Object Example

The following example demonstrates data movement from a  **TextBox** to the Clipboard, from the Clipboard to a **DataObject**, and from a **DataObject** into another **TextBox**. The **GetFromClipboard** method transfers the data from the Clipboard to a **DataObject**. The **Copy** and **GetText** methods are also used.

To use this example, copy this sample code to the Declarations portion of a form. Make sure that the form contains:




- Two  **TextBox** controls named TextBox1 and TextBox2.
    
- A  **CommandButton** named CommandButton1.
    




```vb
Dim MyData as DataObject 
 
Private Sub CommandButton1_Click() 
 'Need to select text before copying it to Clipboard 
 TextBox1.SelStart = 0 
 TextBox1.SelLength = TextBox1.TextLength 
 TextBox1.Copy 
 
 MyData.GetFromClipboard 
 TextBox2.Text = MyData.GetText(1) 
End Sub 
 
Private Sub UserForm_Initialize() 
 Set MyData = New DataObject 
 TextBox1.Text = "Move this data to the " _ 
 &; "Clipboard, to a DataObject, then to " 
 &; "TextBox2!" 
End Sub 
```


