---
title: Name Property Example
keywords: fm20.chm5225156
f1_keywords:
- fm20.chm5225156
ms.prod: office
ms.assetid: d15fecd4-e195-3026-5c7c-5e0780f2f132
ms.date: 06/08/2017
---


# Name Property Example

The following example displays the  **Name** property of each control on a form. This example uses the **Controls** collection to cycle through all the controls placed directly on the Userform.

To use this example, copy this sample code to the Declarations portion of a form. Make sure that the form contains a  **CommandButton** named CommandButton1 and several other controls.



```vb
Private Sub CommandButton1_Click() 
 Dim MyControl As Control 
 
 For Each MyControl In Controls 
 MsgBox "MyControl.Name = " &; MyControl.Name 
 Next 
End Sub
```


