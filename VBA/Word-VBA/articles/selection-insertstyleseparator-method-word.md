---
title: Selection.InsertStyleSeparator Method (Word)
keywords: vbawd10.chm158663676
f1_keywords:
- vbawd10.chm158663676
ms.prod: word
api_name:
- Word.Selection.InsertStyleSeparator
ms.assetid: cbfd7a55-4048-0e16-eeb2-e8d8d167a769
ms.date: 06/08/2017
---


# Selection.InsertStyleSeparator Method (Word)

Inserts a special hidden paragraph mark that allows Microsoft Word to join paragraphs formatted using different paragraph styles, so lead-in headings can be inserted into a table of contents.


## Syntax

 _expression_ . **InsertStyleSeparator**

 _expression_ Required. A variable that represents a **[Selection](selection-object-word.md)** object.


## Example

This example inserts a style separator after every paragraph formatted with the built-in "Heading 4" style.


 **Note**  The paragraph count is inside the Do...Loop because when Word inserts the style separator, the two paragraphs become one paragraph, so the paragraph count for the document changes as the procedure runs.


```vb
Sub InlineHeading() 
 Dim intCount As Integer 
 Dim intParaCount As Integer 
 
 intCount = 1 
 
 With ActiveDocument 
 Do 
 'Look for all paragraphs formatted with "Heading 4" style 
 If .Paragraphs(Index:=intCount).Style = "Heading 4" Then 
 .Paragraphs(Index:=intCount).Range.Select 
 
 'Insert a style separator if paragraph 
 'is formatted with a "Heading 4" style 
 Selection.InsertStyleSeparator 
 End If 
 intCount = intCount + 1 
 intParaCount = .Paragraphs.Count 
 Loop Until intCount = intParaCount 
 
 End With 
End Sub
```


## See also


#### Concepts


[Selection Object](selection-object-word.md)

