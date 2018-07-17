---
title: "How to: Add a Control to a Page"
keywords: olfm10.chm3077152
f1_keywords:
- olfm10.chm3077152
ms.prod: outlook
ms.assetid: 154255a5-7fe7-3397-c239-73a52792c183
ms.date: 06/08/2017
---


# How to: Add a Control to a Page

The following example uses the  **Add** method of the Microsoft Forms 2.0 **Controls** collection to add a control to a form at run time.

To use this example, copy this sample code to the Script Editor of a form. Make sure that the form contains:

- A  **[CommandButton](commandbutton-object-outlook-forms-script.md)** named CommandButton1.
    



```vb
Dim Mycmd 
Sub CommandButton1_Click() 
 
 Set Mycmd = Item.GetInspector.ModifiedFormPages("P.2").Controls.Add("Forms.CommandButton.1") ', CommandButton2, Visible) 
 Mycmd.Left = 18 
 Mycmd.Top = 150 
 Mycmd.Width = 175 
 Mycmd.Height = 20 
 Mycmd.Caption = "This is fun." &; Mycmd.Name 
 
End Sub
```


