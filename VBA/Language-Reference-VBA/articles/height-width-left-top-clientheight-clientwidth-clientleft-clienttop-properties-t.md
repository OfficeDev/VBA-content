---
title: Height, Width, Left, Top, ClientHeight, ClientWidth, ClientLeft, ClientTop Properties, TabStrip, Image Control Example
keywords: fm20.chm5225172
f1_keywords:
- fm20.chm5225172
ms.prod: office
ms.assetid: 26dd7b87-09f1-6f80-0966-913bc39635bd
ms.date: 06/08/2017
---


# Height, Width, Left, Top, ClientHeight, ClientWidth, ClientLeft, ClientTop Properties, TabStrip, Image Control Example

The following example sets the dimensions of an  **Image** to the size of a **TabStrip's** client area when the user clicks a **CommandButton**. This code sample uses the following properties: **Height**, **Left**, **Top**, **Width**, **ClientHeight**, **ClientLeft**, **ClientTop**, and **ClientWidth**.

To use this example, copy this sample code to the Declarations portion of a form. Make sure that the form contains:




- A  **CommandButton** named CommandButton1.
    
- A  **TabStrip** named TabStrip1.
    
- An  **Image** named Image1.
    




```vb
Private Sub UserForm_Initialize() 
 CommandButton1.Caption = "Size Image to Tab Area" 
 CommandButton1.WordWrap = True 
 CommandButton1.AutoSize = True 
End Sub
```




```vb
Private Sub CommandButton1_Click() 
 Image1.ZOrder(fmFront) 
'Place Image in front of TabStrip 
 
'ClientLeft and ClientTop are measured from the edge 
'of the TabStrip, not from the edges of the form 
'containing the TabStrip. 
 Image1.Left = TabStrip1.Left + TabStrip1.ClientLeft 
 Image1.Top = TabStrip1.Top + TabStrip1.ClientTop 
 Image1.Width = TabStrip1.ClientWidth 
 Image1.Height = TabStrip1.ClientHeight 
 
End Sub
```


