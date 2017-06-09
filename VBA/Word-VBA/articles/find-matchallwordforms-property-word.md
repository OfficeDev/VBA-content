---
title: Find.MatchAllWordForms Property (Word)
keywords: vbawd10.chm162529293
f1_keywords:
- vbawd10.chm162529293
ms.prod: word
api_name:
- Word.Find.MatchAllWordForms
ms.assetid: 12244a30-2ddd-8de9-ff74-326c069e656b
ms.date: 06/08/2017
---


# Find.MatchAllWordForms Property (Word)

 **True** if all forms of the text to find are found by the find operation (for instance, if the text to find is "sit," "sat" and "sitting" are found as well). Read/write **Boolean** .


## Syntax

 _expression_ . **MatchAllWordForms**

 _expression_ An expression that returns a **[Find](find-object-word.md)** object.


## Remarks

Use the  **[Text](find-text-property-word.md)** property of the **Find** object or use the FindText argument with the **[Execute](find-execute-method-word.md)** method to specify the text to be located in a document.


## Example

This example selects the next form of the word "sit" found in the selection or displays a message box if a form of "sit" isn't found.


```vb
With Selection.Find 
 .MatchAllWordForms = True 
 .Text = "sit" 
 .Execute Format:=False 
 If .Found = False Then MsgBox "Not Found" 
End With
```


## See also


#### Concepts


[Find Object](find-object-word.md)

