---
title: ProtectedViewWindow.Caption Property (Word)
keywords: vbawd10.chm231735296
f1_keywords:
- vbawd10.chm231735296
ms.prod: word
api_name:
- Word.ProtectedViewWindow.Caption
ms.assetid: ec8d2b22-34b6-2685-6ab5-74eb48b1dfb0
ms.date: 06/08/2017
---


# ProtectedViewWindow.Caption Property (Word)

Returns or sets the caption text that is displayed in the title bar of the document or protected view window. Read/write  **String** .


## Syntax

 _expression_ . **Caption**

 _expression_ An expression that returns a **[ProtectedViewWindow](protectedviewwindow-object-word.md)** object.


## Remarks

To change the caption of the protected view window to the default text, set this property to an empty string ("").


## Example

The following code example displays the caption for the active protected view window.


```vb
MsgBox "The caption for the active protected " &; _ 
 "view window is: " &; ActiveProtectedViewWindow.Caption 

```

The following code example changes the caption for the active protected view window.




```vb
ActiveProtectedViewWindow.Caption = Application.UserName &; "'s copy of Word" 

```


## See also


#### Concepts


[ProtectedViewWindow Object](protectedviewwindow-object-word.md)

