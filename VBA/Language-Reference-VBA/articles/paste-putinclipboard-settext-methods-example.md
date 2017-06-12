---
title: Paste, PutInClipboard, SetText Methods Example
keywords: fm20.chm5225164
f1_keywords:
- fm20.chm5225164
ms.prod: office
ms.assetid: d7045eb8-3b79-a490-91a8-b6f5369bbf8c
ms.date: 06/08/2017
---


# Paste, PutInClipboard, SetText Methods Example

The following example demonstrates data movement from a  **TextBox** to a **DataObject**, from a **DataObject** to the Clipboard, and from the Clipboard to another **TextBox**. The **PutInClipboard** method transfers the data from a **DataObject** to the Clipboard. The **SetText** and **Paste** methods are also used.

To use this example, copy this sample code to the Declarations portion of a form. Make sure that the form contains:




- Two  **TextBox** controls named TextBox1 and TextBox2.
    
- A  **CommandButton** named CommandButton1.
    




```vb
Dim MyData As DataObject 
 
Private Sub CommandButton1_Click() 
 Set MyData = New DataObject 
 
 MyData.SetText TextBox1.Text 
 MyData.PutInClipboard 
 
 TextBox2.Paste 
End Sub 
 
Private Sub UserForm_Initialize() 
 TextBox1.Text = "Move this data to a " _ 
 &; "DataObject, to the Clipboard, then to " _ 
 &; "TextBox2!" 
End Sub
```


