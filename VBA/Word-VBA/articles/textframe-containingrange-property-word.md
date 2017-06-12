---
title: TextFrame.ContainingRange Property (Word)
keywords: vbawd10.chm162661354
f1_keywords:
- vbawd10.chm162661354
ms.prod: word
api_name:
- Word.TextFrame.ContainingRange
ms.assetid: c6e4cf7e-1f4a-232f-1e55-5cbb4537df8a
ms.date: 06/08/2017
---


# TextFrame.ContainingRange Property (Word)

Returns a  **[Range](range-object-word.md)** object that represents the entire story in a series of shapes with linked text frames that the specified text frame belongs to. Read-only.


## Syntax

 _expression_ . **ContainingRange**

 _expression_ A variable that represents a **[TextFrame](textframe-object-word.md)** object.


## Example

This example checks the spelling in TextBox 1 and any other text in text frames that are linked to TextBox 1.


```vb
Dim rngStory As Range 
 
Set rngStory = ActiveDocument.Shapes("TextBox 1") _ 
 .TextFrame.ContainingRange 
 
rngStory.CheckSpelling
```


## See also


#### Concepts


[TextFrame Object](textframe-object-word.md)

