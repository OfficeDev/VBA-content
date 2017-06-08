---
title: Style.Description Property (Word)
keywords: vbawd10.chm153878530
f1_keywords:
- vbawd10.chm153878530
ms.prod: word
api_name:
- Word.Style.Description
ms.assetid: fec1fa70-7080-e159-b20c-1a389cbaf903
ms.date: 06/08/2017
---


# Style.Description Property (Word)

Returns the description of the specified style. Read-only  **String** .


## Syntax

 _expression_ . **Description**

 _expression_ A variable that represents a **[Style](style-object-word.md)** object.


## Remarks

A typical example of a descirption for a style might be "Normal + Font: Arial, 12 pt, Bold, Italic, Space Before 12 pt After 3 pt, KeepWithNext, Level 2."


## Example

This example creates a new document and inserts a tab-delimited list of the active document's styles and their descriptions.


```vb
Dim docActive As Document 
Dim docNew As Document 
Dim styleLoop As Style 
 
Set docActive = ActiveDocument 
Set docNew = Documents.Add 
 
For Each styleLoop In docActive.Styles 
 With docNew.Range 
 .InsertAfter Text:=styleLoop.NameLocal &; Chr(9) _ 
 &; styleLoop.Description 
 .InsertParagraphAfter 
 End With 
Next styleLoop
```


## See also


#### Concepts


[Style Object](style-object-word.md)

