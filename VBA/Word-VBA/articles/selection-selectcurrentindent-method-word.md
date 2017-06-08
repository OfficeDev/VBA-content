---
title: Selection.SelectCurrentIndent Method (Word)
keywords: vbawd10.chm158663176
f1_keywords:
- vbawd10.chm158663176
ms.prod: word
api_name:
- Word.Selection.SelectCurrentIndent
ms.assetid: 3a71080e-935c-fc3c-40b9-e82acf9d28cc
ms.date: 06/08/2017
---


# Selection.SelectCurrentIndent Method (Word)

Extends the selection forward until text with different left or right paragraph indents is encountered.


## Syntax

 _expression_ . **SelectCurrentIndent**

 _expression_ Required. A variable that represents a **[Selection](selection-object-word.md)** object.


## Example

This example jumps to the beginning of the first paragraph in the document that has different indents than the first paragraph in the active document.


```vb
With Selection 
 .HomeKey Unit:=wdStory, Extend:=wdMove 
 .SelectCurrentIndent 
 .Collapse Direction:=wdCollapseEnd 
End With
```

This example determines whether all the paragraphs in the active document are formatted with the same left and right indents and then displays a message box indicating the result.




```vb
With Selection 
 .HomeKey Unit:=wdStory, Extend:=wdMove 
 .SelectCurrentIndent 
 .Collapse Direction:=wdCollapseEnd 
End With 
If Selection.End = ActiveDocument.Content.End - 1 Then 
 MsgBox "All paragraphs share the same left " _ 
 &; "and right indents." 
Else 
 MsgBox "Not all paragraphs share the same left " _ 
 &; "and right indents." 
End If
```


## See also


#### Concepts


[Selection Object](selection-object-word.md)

