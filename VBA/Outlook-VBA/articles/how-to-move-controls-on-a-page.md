---
title: "How to: Move Controls on a Page"
keywords: olfm10.chm3077219
f1_keywords:
- olfm10.chm3077219
ms.prod: outlook
ms.assetid: d50e7b95-016d-9ee7-533a-4a101e2316eb
ms.date: 06/08/2017
---


# How to: Move Controls on a Page

The following example demonstrates moving all the controls on a form by using the  **Move** method with the Microsoft Forms 2.0 **Controls** collection. The user clicks on the **[CommandButton](commandbutton-object-outlook-forms-script.md)** to move the controls.

To use this example, copy this sample code to the Script Editor of a form. Make sure that the form contains a  **CommandButton** named CommandButton1 and several other controls.



```vb
Sub CommandButton1_Click() 
 Set Controls = Item.GetInspector.ModifiedFormPages("P.2").Controls 
 'Move each control on the form right 25 points and up 25 points. 
 Controls.Move 25, -25 
End Sub
```


