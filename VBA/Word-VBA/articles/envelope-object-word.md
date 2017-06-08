---
title: Envelope Object (Word)
keywords: vbawd10.chm2328
f1_keywords:
- vbawd10.chm2328
ms.prod: word
api_name:
- Word.Envelope
ms.assetid: 03664453-f7fb-f76a-ea60-37e72b53e17c
ms.date: 06/08/2017
---


# Envelope Object (Word)

Represents an envelope attached to a document.


## Remarks

Use the  **[Envelope](document-envelope-property-word.md)** property to return the **Envelope** object. The following example adds an envelope to a new document and sets the distance between the top of the envelope and the address to 2.25 inches.


```vb
Set myDoc = Documents.Add 
addr = "Michael Matey" &; vbCr &; "123 Skye St." _ 
 &; vbCr &; "Redmond, WA 98107" 
retaddr = "Cora Edmonds" &; vbCr &; "456 Erde Lane" &; vbCr _ 
 &; "Redmond, WA 98107" 
With myDoc.Envelope 
 .Insert Address:=addr, ReturnAddress:=retaddr 
 .AddressFromTop = InchesToPoints(2.25) 
End With
```

Remarks

The  **Envelope** object is available regardless of whether an envelope has been added to the specified document. However, an error occurs if you use one of the following properties when an envelope has not been added to the document: **[Address](envelope-address-property-word.md)** , **[AddressFromLeft](envelope-addressfromleft-property-word.md)** , **[AddressFromTop](envelope-addressfromtop-property-word.md)** , **[FeedSource](envelope-feedsource-property-word.md)** , **[ReturnAddress](envelope-returnaddress-property-word.md)** , **[ReturnAddressFromLeft](envelope-returnaddressfromleft-property-word.md)** , **[ReturnAddressFromTop](envelope-returnaddressfromtop-property-word.md)** , and **[UpdateDocument](envelope-updatedocument-method-word.md)** .

The following example demonstrates how to use the  **On Error GoTo** statement to trap the error that occurs if an envelope has not been added to the active document. If, however, an envelope has been added to the document, the recipient address is displayed.




```vb
On Error GoTo ErrorHandler 
MsgBox ActiveDocument.Envelope.Address 
ErrorHandler: 
If Err = 5852 Then MsgBox _ 
 "Envelope is not in the specified document"
```

Use the  **Insert** method to add an envelope to the specified document. Use the **PrintOut** method to set the properties of an envelope and print it without adding it to the document.


 **Note**  There is no Envelopes collection; each  **Document** object contains only one **Envelope** object.


## See also


#### Other resources



[Word Object Model Reference](http://msdn.microsoft.com/library/be452561-b436-bb9b-6f94-3faa9a74a6fd%28Office.15%29.aspx)

