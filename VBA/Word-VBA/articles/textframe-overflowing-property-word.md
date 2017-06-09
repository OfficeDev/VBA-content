---
title: TextFrame.Overflowing Property (Word)
keywords: vbawd10.chm162665355
f1_keywords:
- vbawd10.chm162665355
ms.prod: word
api_name:
- Word.TextFrame.Overflowing
ms.assetid: 299020e0-0c26-e5cb-c47c-2aa3651aac36
ms.date: 06/08/2017
---


# TextFrame.Overflowing Property (Word)

 **True** if the text inside the specified text frame doesn't all fit within the frame. Read-only **Boolean** .


## Syntax

 _expression_ . **Overflowing**

 _expression_ An expression that returns a **[TextFrame](textframe-object-word.md)** object.


## Example

This example checks to see whether the text in MyTextBox is overflowing its text frame. If so, the example adds another text box and links the two text boxes so that the text flows into the next one.


```vb
Set myTBox = ActiveDocument.Shapes("MyTextBox") 
If myTBox.TextFrame.Overflowing = True Then 
 Set nextTBox = ActiveDocument.Shapes. _ 
 AddTextbox(msoTextOrientationHorizontal, 72, 72, 100, 200) 
 MyTBox.TextFrame.Next = nextTBox.TextFrame 
End If
```


## See also


#### Concepts


[TextFrame Object](textframe-object-word.md)

