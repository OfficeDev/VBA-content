---
title: Envelope.Insert Method (Word)
keywords: vbawd10.chm152567913
f1_keywords:
- vbawd10.chm152567913
ms.prod: word
api_name:
- Word.Envelope.Insert
ms.assetid: 6fd42ed0-f8d0-d2be-175d-345f1367de61
ms.date: 06/08/2017
---


# Envelope.Insert Method (Word)

Inserts an envelope as a separate section at the beginning of the specified document.


## Syntax

 _expression_ . **Insert**( **_ExtractAddress_** , **_Address_** , **_AutoText_** , **_OmitReturnAddress_** , **_ReturnAddress_** , **_ReturnAutoText_** , **_PrintBarCode_** , **_PrintFIMA_** , **_Size_** , **_Height_** , **_Width_** , **_FeedSource_** , **_AddressFromLeft_** , **_AddressFromTop_** , **_ReturnAddressFromLeft_** , **_ReturnAddressFromTop_** , **_DefaultFaceUp_** , **_DefaultOrientation_** , **_PrintEPostage_** , **_Vertical_** , **_RecipientNamefromLeft_** , **_RecipientNamefromTop_** , **_RecipientPostalfromLeft_** , **_RecipientPostalfromTop_** , **_SenderNamefromLeft_** , **_SenderNamefromTop_** , **_SenderPostalfromLeft_** , **_SenderPostalfromTop_** )

 _expression_ Required. A variable that represents an **[Envelope](envelope-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _ExtractAddress_|Optional| **Variant**| **True** to use the text marked by the EnvelopeAddress bookmark (a user-defined bookmark) as the recipient's address.|
| _Address_|Optional| **Variant**|A string that specifies the recipient's address (ignored if ExtractAddress is  **True** ).|
| _AutoText_|Optional| **Variant**|A string that specifies an AutoText entry to use for the address. If specified, Address is ignored.|
| _OmitReturnAddress_|Optional| **Variant**| **True** to not insert a return address.|
| _ReturnAddress_|Optional| **Variant**|A string that specifies the return address.|
| _ReturnAutoText_|Optional| **Variant**|A string that specifies an AutoText entry to use for the return address. If specified, ReturnAddress is ignored.|
| _PrintBarCode_|Optional| **Variant**| **True** to add a POSTNET bar code. For U.S. mail only.|
| _PrintFIMA_|Optional| **Variant**| **True** to add a Facing Identification Mark (FIMA) for use in presorting courtesy reply mail. For U.S. mail only.|
| _Size_|Optional| **Variant**|A string that specifies the envelope size. The string must match one of the sizes listed in the  **Envelope size** box in the **Envelope Options** dialog box (for example, "Size 10" or "C4").|
| _Height_|Optional| **Variant**|The height of the envelope, measured in points, when the Size argument is set to "Custom size."|
| _Width_|Optional| **Variant**|The width of the envelope, measured in points, when the Size argument is set to "Custom size."|
| _FeedSource_|Optional| **Variant**| **True** to use the **FeedSource** property of the **Envelope** object to specify which paper tray to use when printing the envelope.|
| _AddressFromLeft_|Optional| **Variant**|The distance, measured in points, between the left edge of the envelope and the recipient's address.|
| _AddressFromTop_|Optional| **Variant**|The distance, measured in points, between the top edge of the envelope and the recipient's address.|
| _ReturnAddressFromLeft_|Optional| **Variant**|The distance, measured in points, between the left edge of the envelope and the return address.|
| _ReturnAddressFromTop_|Optional| **Variant**|The distance, measured in points, between the top edge of the envelope and the return address.|
| _DefaultFaceUp_|Optional| **Variant**| **True** to print the envelope face up, **False** to print it face down.|
| _DefaultOrientation_|Optional| **Variant**|The orientation for the envelope. Can be any  **WdEnvelopeOrientation** constant.|
| _PrintEPostage_|Optional| **Variant**| **True** to insert postage from an Internet postage vendor.|
| _Vertical_|Optional| **Variant**| **True** to print vertical text on the envelope. Used for Asian envelopes. Default is **False** .|
| _RecipientNamefromLeft_|Optional| **Variant**|Position of the recipient's name, measured in points from the left edge of the envelope. Used for Asian envelopes.|
| _RecipientNamefromTop_|Optional| **Variant**|Position of the recipient's name, measured in points from the top edge of the envelope. Used for Asian envelopes.|
| _RecipientPostalfromLeft_|Optional| **Variant**|Position of the recipient's postal code, measured in points from the left edge of the envelope. Used for Asian envelopes.|
| _RecipientPostalfromTop_|Optional| **Variant**|Position of the recipient's postal code, measured in points from the top edge of the envelope. Used for Asian envelopes.|
| _SenderNamefromLeft_|Optional| **Variant**|Position of the sender's name, measured in points from the left edge of the envelope. Used for Asian envelopes.|
| _SenderNamefromTop_|Optional| **Variant**|Position of the sender's name, measured in points from the top edge of the envelope. Used for Asian envelopes.|
| _SenderPostalfromLeft_|Optional| **Variant**|Position of the sender's postal code, measured in points from the left edge of the envelope. Used for Asian envelopes.|
| _SenderPostalfromTop_|Optional| **Variant**|Position of the sender's postal code, measured in points from the top edge of the envelope. Used for Asian envelopes.|

## Example

This example adds a Size 10 envelope to the active document by using the addresses stored in the  _strAddr_ and _strReturnAddr_ variables.


```vb
Sub InsertEnvelope() 
 Dim strAddr As String 
 Dim strReturnAddr As String 
 strAddr = "Max Benson" &; vbCr &; "123 Skye St." _ 
 &; vbCr &; "OurTown, WA 98107" 
 strReturnAddr = "Paul Borm" &; vbCr &; "456 Erde Lane" _ 
 &; vbCr &; "OurTown, WA 98107" 
 ActiveDocument.Envelope.Insert Address:=strAddr, _ 
 ReturnAddress:=strReturnAddr, Size:="Size 10" 
End Sub
```


## See also


#### Concepts


[Envelope Object](envelope-object-word.md)

