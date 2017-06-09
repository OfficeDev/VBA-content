---
title: Document.Range Method (Word)
keywords: vbawd10.chm158009296
f1_keywords:
- vbawd10.chm158009296
ms.prod: word
api_name:
- Word.Document.Range
ms.assetid: 7dd33ac8-38f6-925d-a511-8677ca6437e6
ms.date: 06/08/2017
---


# Document.Range Method (Word)

Returns a  **Range** object by using the specified starting and ending character positions.


## Syntax

 _expression_ . **Range**( **_Start_** , **_End_** )

 _expression_ Required. A variable that represents a **[Document](document-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Start_|Optional| **Variant**|The starting character position.|
| _End_|Optional| **Variant**|The ending character position.|

### Return Value

Range


## Example

This example applies bold formatting to the first 10 characters in the active document.


```vb
Sub DocumentRange() 
 ActiveDocument.Range(Start:=0, End:=10).Bold = True 
End Sub
```

This example creates a range that starts at the beginning of the active document and ends at the cursor position, and then it changes all characters within that range to uppercase.




```vb
Sub DocumentRange2() 
 Dim r As Range 
 Set r = ActiveDocument.Range(Start:=0, End:=Selection.End) 
 r.Case = wdUpperCase 
End Sub
```

This example creates and sets the variable  _myRange_ to paragraphs three through six in the active document, and then it right-aligns the paragraphs in the range.




```vb
Sub DocumentRange3() 
 Dim aDoc As Document 
 Dim myRange As Range 
 Set aDoc = ActiveDocument 
 If aDoc.Paragraphs.Count >= 6 Then 
 Set myRange = aDoc.Range(aDoc.Paragraphs(2).Range.Start, _ 
 aDoc.Paragraphs(4).Range.End) 
 myRange.Paragraphs.Alignment = wdAlignParagraphRight 
 End If 
End Sub
```


## See also


#### Concepts


[Document Object](document-object-word.md)

