---
title: Move Method Example
keywords: fm20.chm5225194
f1_keywords:
- fm20.chm5225194
ms.prod: office
ms.assetid: c5444339-b059-9b55-a3a4-9e5b4e2573f6
ms.date: 06/08/2017
---


# Move Method Example

The following example demonstrates moving all the controls on a form by using the  **Move** method with the **Controls** collection. The user clicks on the **CommandButton** to move the controls.

To use this example, copy this sample code to the Declarations portion of a form. Make sure that the form contains a  **CommandButton** named CommandButton1 and several other controls.



```vb
Private Sub CommandButton1_Click() 
 'Move each control on the form right 25 points 
 'and up 25 points. 
Controls.Move 25, -25 
End Sub
```


