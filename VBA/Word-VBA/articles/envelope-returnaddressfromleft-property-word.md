---
title: Envelope.ReturnAddressFromLeft Property (Word)
keywords: vbawd10.chm152567823
f1_keywords:
- vbawd10.chm152567823
ms.prod: word
api_name:
- Word.Envelope.ReturnAddressFromLeft
ms.assetid: ab0a068b-0b66-481b-ca07-25bb17e2abcf
ms.date: 06/08/2017
---


# Envelope.ReturnAddressFromLeft Property (Word)

Returns or sets the distance (in points) between the left edge of the envelope and the return address. Read/write  **Single** .


## Syntax

 _expression_ . **ReturnAddressFromLeft**

 _expression_ An expression that returns an **[Envelope](envelope-object-word.md)** object.


## Remarks

If you use this property before an envelope has been added to the document, an error occurs.


## Example

This example creates a new document and adds an envelope with a predefined delivery address and return address. The example then sets the distance between the left edge of the envelope and the return address to 0.75 inch.


```vb
addr = "Karin Gallagher" &; vbCr &; "123 Skye St." _ 
 &; vbCr &; "Our Town, WA 98004" 
retaddr = "Don Funk" &; vbCr &; "123 Main" _ 
 &; vbCr &; "Other Town, WA 98040" 
With Documents.Add.Envelope 
 .Insert Address:=addr, ReturnAddress:=retaddr 
 .ReturnAddressFromLeft = InchesToPoints(0.75) 
End With 
ActiveDocument.ActiveWindow.View.Type = wdPrintView
```


## See also


#### Concepts


[Envelope Object](envelope-object-word.md)

