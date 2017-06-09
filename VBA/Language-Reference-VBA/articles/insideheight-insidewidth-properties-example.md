---
title: InsideHeight, InsideWidth Properties Example
keywords: fm20.chm5225139
f1_keywords:
- fm20.chm5225139
ms.prod: office
ms.assetid: 5b6c7176-0838-33da-1111-9591f961641e
ms.date: 06/08/2017
---


# InsideHeight, InsideWidth Properties Example

The following example uses the  **InsideHeight** and **InsideWidth** properties to resize a **CommandButton**. The user clicks the **CommandButton** to resize it.


 **Note**   **InsideHeight** and **InsideWidth** are read-only properties.


To use this example, copy this sample code to the Declarations portion of a form. Make sure that the form contains:



- A  **CommandButton** named CommandButton1.
    




```vb
Dim Resize As Single 
 
Private Sub UserForm_Initialize() 
 Resize = 0.75 
 CommandButton1.Caption = "Resize Button" 
 
End Sub 
 
Private Sub CommandButton1_Click() 
 CommandButton1.Move 10, 10, _ 
 UserForm1.InsideWidth * Resize, _ 
 UserForm1.InsideHeight * Resize 
 CommandButton1.Caption = "Button resized " _ 
 &; "using InsideHeight and InsideWidth!" 
End Sub
```


