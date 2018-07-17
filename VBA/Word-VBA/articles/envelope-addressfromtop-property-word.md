---
title: Envelope.AddressFromTop Property (Word)
keywords: vbawd10.chm152567822
f1_keywords:
- vbawd10.chm152567822
ms.prod: word
api_name:
- Word.Envelope.AddressFromTop
ms.assetid: 425eb517-85af-68e2-951b-66282b813e9b
ms.date: 06/08/2017
---


# Envelope.AddressFromTop Property (Word)

Returns or sets the distance (in points) between the top edge of the envelope and the delivery address. Read/write  **Single** .


## Syntax

 _expression_ . **AddressFromTop**

 _expression_ A variable that represents a **[Envelope](envelope-object-word.md)** object.


## Remarks

If you use this property before an envelope has been added to the document, an error occurs.


## Example

This example creates a new document and adds an envelope with a predefined delivery address and return address. The example then sets the distance between the top edge of the envelope and the delivery address to 1.75 inches and sets the distance between the left edge of the envelope and the delivery address is set to 3.75 inches.


```vb
Dim strAddress As String 
Dim strReturn As String 

```


```vb
strAddress = "Michael Bunney" &; vbCr &; "123 Skye St." &; vbCr _ 
 &; "Our Town, WA 98040" 
strReturn = "Kate Dresen" &; vbCr &; "123 Main" &; vbCr _ 
 &; "Other Town, WA 98040" 
 
With Documents.Add.Envelope 
 .Insert Address:=strAddress, ReturnAddress:=strReturn 
 .AddressFromTop = InchesToPoints(1.75) 
 .AddressFromLeft = InchesToPoints(3.75) 
End With 
 
ActiveDocument.ActiveWindow.View.Type = wdPrintView
```


## See also


#### Concepts


[Envelope Object](envelope-object-word.md)

