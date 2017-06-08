---
title: "How to: Cut Text from One Text Box and Paste it on Another"
keywords: olfm10.chm3077171
f1_keywords:
- olfm10.chm3077171
ms.prod: outlook
ms.assetid: 33339831-9567-6910-f596-6a9a398886e8
ms.date: 06/08/2017
---


# How to: Cut Text from One Text Box and Paste it on Another

The following example uses the  **[Cut](textbox-cut-method-outlook-forms-script.md)** and **[Paste](textbox-paste-method-outlook-forms-script.md)** methods to cut text from one **[TextBox](textbox-object-outlook-forms-script.md)** and paste it on another **TextBox**.

To use this example, copy this sample code to the Script Editor of a form. To run the code you need to open the form so the  **Open** event will activate. Make sure that the form contains:

- Two  **TextBox** controls named TextBox1 and TextBox2.
    
- A  **[CommandButton](commandbutton-object-outlook-forms-script.md)** named CommandButton1.
    



```vb
Dim TextBox1 
Dim TextBox2 
Dim CommandButton1 
 
Sub Item_Open() 
 Set TextBox1 = Item.GetInspector.ModifiedFormPages.Item("P.2").Controls("TextBox1") 
 Set TextBox2 = Item.GetInspector.ModifiedFormPages.Item("P.2").Controls("TextBox2") 
 Set CommandButton1 = Item.GetInspector.ModifiedFormPages.Item("P.2").Controls("CommandButton1") 
 
 TextBox1.Text = "From TextBox1!" 
 TextBox2.Text = "Hello " 
 
 CommandButton1.Caption = "Cut and Paste" 
 CommandButton1.AutoSize = True 
End Sub 
 
Sub CommandButton1_Click() 
 TextBox2.SelStart = 0 
 TextBox2.SelLength = TextBox2.TextLength 
 TextBox2.Cut 
 
 TextBox1.SetFocus 
 TextBox1.SelStart = 0 
 
 TextBox1.Paste 
 TextBox2.SelStart = 0 
End Sub
```


