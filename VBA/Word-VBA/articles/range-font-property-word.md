---
title: Range.Font Property (Word)
keywords: vbawd10.chm157155333
f1_keywords:
- vbawd10.chm157155333
ms.prod: word
api_name:
- Word.Range.Font
ms.assetid: 7582a7ed-0f16-e8f3-73f7-5d7b91193679
ms.date: 06/08/2017
---


# Range.Font Property (Word)

Returns or sets a  **[Font](font-object-word.md)** object that represents the character formatting of the specified object. Read/write **Font** .


## Syntax

 _expression_ . **Font**

 _expression_ A variable that represents a **[Range](range-object-word.md)** object.


## Remarks

To set this property, specify an expression that returns a  **[Font](font-object-word.md)** object.


## Example

This example removes bold formatting from the Heading 1 style in the active document.


```vb
ActiveDocument.Styles(wdStyleHeading1).Font.Bold = False
```

This example switches the font of the second paragraph in the active document between Arial and Times New Roman.




```vb
Set myRange = ActiveDocument.Paragraphs(2).Range 
If myRange.Font.Name = "Times New Roman" Then 
 myRange.Font.Name = "Arial" 
Else 
 myRange.Font.Name = "Times New Roman" 
End If
```


## See also


#### Concepts


[Range Object](range-object-word.md)

