---
title: "How to: Display the Name of Each Control on a Page"
keywords: olfm10.chm3077222
f1_keywords:
- olfm10.chm3077222
ms.prod: outlook
ms.assetid: 109f6397-5b7f-bd8d-0ef5-ed0ba770bc5b
ms.date: 06/08/2017
---


# How to: Display the Name of Each Control on a Page

The following example displays the  **Name** property of each control on a form. This example uses the Microsoft Forms 2.0 **Controls**collection to cycle through all the controls placed directly on the User form.

To use this example, copy this sample code to the Script Editor of a form. Make sure that the form contains a  **[CommandButton](commandbutton-object-outlook-forms-script.md)** named CommandButton1 and several other controls.



```vb
Sub CommandButton1_Click() 
 Set Controls = Item.GetInspector.ModifiedFormPages.Item("P.2").Controls 
 For i = 0 to Controls.Count - 1 
 MsgBox "MyControl.Name = " &; Controls.Item(i).Name 
 Next 
End Sub
```


