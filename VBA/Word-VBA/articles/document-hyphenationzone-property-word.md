---
title: Document.HyphenationZone Property (Word)
keywords: vbawd10.chm158007309
f1_keywords:
- vbawd10.chm158007309
ms.prod: word
api_name:
- Word.Document.HyphenationZone
ms.assetid: 30ea2a99-a8f5-10f4-58f9-48533bf3ec00
ms.date: 06/08/2017
---


# Document.HyphenationZone Property (Word)

Returns or sets the width of the hyphenation zone, in points. Read/write  **Long** .


## Syntax

 _expression_ . **HyphenationZone**

 _expression_ A variable that represents a **[Document](document-object-word.md)** object.


## Remarks

The hyphenation zone is the maximum amount of space that Microsoft Word leaves between the end of the last word in a line and the right margin.


## Example


 **Note**  Unless Word is in compatibility mode,  **HyphenationZone** always returns 99999999.

This example enables automatic hyphenation for MyReport.doc. The hyphenation zone is set to 36 points (0.5 inch).


```vb
With Documents("MyReport.doc") 
 .AutoHyphenation = True 
 .HyphenationZone = 36 
End With
```

This example sets the hyphenation zone to 0.25 inch (18 points) and then starts manual hyphenation of the active document.




```vb
With ActiveDocument 
 .HyphenationZone = InchesToPoints(0.25) 
 .ManualHyphenation 
End With
```


## See also


#### Concepts


[Document Object](document-object-word.md)

