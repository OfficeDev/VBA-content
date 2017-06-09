---
title: TextFrame.HasText Property (Word)
keywords: vbawd10.chm162665360
f1_keywords:
- vbawd10.chm162665360
ms.prod: word
api_name:
- Word.TextFrame.HasText
ms.assetid: eb3d99ed-b65f-e0d3-b18f-388cec86bd3d
ms.date: 06/08/2017
---


# TextFrame.HasText Property (Word)

 **True** if the specified shape has text associated with it. Read-only **Boolean** .


## Syntax

 _expression_ . **HasText**

 _expression_ A variable that represents a **[TextFrame](textframe-object-word.md)** object.


## Example

If the second shape on the active document contains text, this example displays a message if the text overflows its frame.


```vb
Dim docActive As Document 
 
Set docActive = ActiveDocument 
With docActive.Shapes(2).TextFrame 
 If .HasText = True Then 
 If .Overflowing = True Then 
 Msgbox "Text overflows the frame." 
 End If 
 End If 
End With
```


## See also


#### Concepts


[TextFrame Object](textframe-object-word.md)

