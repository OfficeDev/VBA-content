---
title: Envelope.AddressFromLeft Property (Word)
keywords: vbawd10.chm152567821
f1_keywords:
- vbawd10.chm152567821
ms.prod: word
api_name:
- Word.Envelope.AddressFromLeft
ms.assetid: 452734c0-fa41-8c90-2478-ecbd5731d393
ms.date: 06/08/2017
---


# Envelope.AddressFromLeft Property (Word)

Returns or sets the distance (in points) between the left edge of the envelope and the delivery address. Read/write  **Single** .


## Syntax

 _expression_ . **AddressFromLeft**

 _expression_ A variable that represents a **[Envelope](envelope-object-word.md)** object.


## Remarks

If you use this property before an envelope has been added to the document, an error occurs.


## Example

This example creates a new document and adds an envelope with a predefined delivery address and return address. The example then sets the distance between the left edge of the envelope and the delivery address to 3.75 inches.


```vb
Dim strAddress As String 
Dim strReturn As String 
 
strAddress = "James Allard" &; vbCr &; "123 Skye St." &; vbCr _ 
 &; "Our Town, WA 98004" 
strReturn = "Rich Andrews" &; vbCr &; "123 Main" &; vbCr _ 
 &; "Other Town, WA 98004" 
 
With Documents.Add.Envelope 
 .Insert Address:=strAddress, ReturnAddress:=strReturn 
 .AddressFromLeft = InchesToPoints(3.75) 
End With 
ActiveDocument.ActiveWindow.View.Type = wdPrintView
```


## See also


#### Concepts


[Envelope Object](envelope-object-word.md)

