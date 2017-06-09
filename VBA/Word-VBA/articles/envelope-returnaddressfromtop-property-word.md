---
title: Envelope.ReturnAddressFromTop Property (Word)
keywords: vbawd10.chm152567824
f1_keywords:
- vbawd10.chm152567824
ms.prod: word
api_name:
- Word.Envelope.ReturnAddressFromTop
ms.assetid: 14738afb-17ab-c1d3-8de5-4fb7a34fa478
ms.date: 06/08/2017
---


# Envelope.ReturnAddressFromTop Property (Word)

Returns or sets the distance (in points) between the top edge of the envelope and the return address. Read/write  **Single** .


## Syntax

 _expression_ . **ReturnAddressFromTop**

 _expression_ An expression that returns an **[Envelope](envelope-object-word.md)** object.


## Remarks

If you use this property before an envelope has been added to the document, an error occurs.


## Example

This example creates a new document and adds an envelope with a predefined delivery address and return address. The example then sets the distance between the top edge of the envelope and the return address to 0.5 inch and sets the distance between the left edge of the envelope and the return address to 0.75 inch.


```vb
addr = "Eric Lang" &; vbCr &; "123 Main" _ 
 &; vbCr &; "Seattle, WA 98040" 
retaddr = "Nate Sun" &; vbCr &; "123 Main" _ 
 &; vbCr &; "Bellevue, WA 98004" 
With Documents.Add.Envelope 
 .Insert Address:=addr, ReturnAddress:=retaddr 
 .ReturnAddressFromTop = InchesToPoints(0.5) 
 .ReturnAddressFromLeft = InchesToPoints(0.75) 
End With
```


## See also


#### Concepts


[Envelope Object](envelope-object-word.md)

