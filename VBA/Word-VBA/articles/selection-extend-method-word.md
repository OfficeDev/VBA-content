---
title: Selection.Extend Method (Word)
keywords: vbawd10.chm158662956
f1_keywords:
- vbawd10.chm158662956
ms.prod: word
api_name:
- Word.Selection.Extend
ms.assetid: 7f9108a1-9b23-bc45-61f5-49aca9979932
ms.date: 06/08/2017
---


# Selection.Extend Method (Word)

Turns on extend mode, or if extend mode is already on, extends the selection to the next larger unit of text.


## Syntax

 _expression_ . **Extend**( **_Character_** )

 _expression_ Required. A variable that represents a **[Selection](selection-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Character_|Optional| **Variant**|The character through which the selection is extended. This argument is case sensitive and must evaluate to a  **String** or an error occurs. Also, if the value of this argument is longer than a single character, Microsoft Word ignores the command entirely.|

## Remarks

Using this method sets the  **ExtendMode** property to **True** if it is not already.

 The progression of selected units of text is as follows: word, sentence, paragraph, section, entire document. If Character is specified, this method extends the selection forward through the next instance of the specified character. The selection is extended by moving the active end of the selection.


## Example

This example collapses the current selection to an insertion point and then selects the current sentence.


```vb
With Selection 
 ' Collapse current selection to insertion point. 
 .Collapse 
 ' Turn extend mode on. 
 .Extend 
 ' Extend selection to word. 
 .Extend 
 ' Extend selection to sentence. 
 .Extend 
End With
```

Here is an example that accomplishes the same task without the  **Extend** method.




```vb
With Selection 
 ' Collapse current selection. 
 .Collapse 
 ' Expand selection to current sentence. 
 .Expand Unit:=wdSentence 
End With
```

This example makes the end of the selection active and extends the selection through the next instance of a capital "R".




```vb
With Selection 
 .StartIsActive = False 
 .Extend Character:="R" 
End Wit
```


## See also


#### Concepts


[Selection Object](selection-object-word.md)

