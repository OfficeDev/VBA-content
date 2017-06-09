---
title: Envelope.PrintOut Method (Word)
keywords: vbawd10.chm152567914
f1_keywords:
- vbawd10.chm152567914
ms.prod: word
api_name:
- Word.Envelope.PrintOut
ms.assetid: 68d8d60a-f07a-1371-e9cc-1d08118e5295
ms.date: 06/08/2017
---


# Envelope.PrintOut Method (Word)

Prints an envelope without adding the envelope to the active document.


## Syntax

 _expression_ . **PrintOut**( **_ExtractAddress_** , **_Address_** , **_AutoText_** , **_OmitReturnAddress_** , **_ReturnAddress_** , **_ReturnAutoText_** , **_PrintBarCode_** , **_PrintFIMA_** , **_Size_** , **_Height_** , **_Width_** , **_FeedSource_** , **_AddressFromLeft_** , **_AddressFromTop_** , **_ReturnAddressFromLeft_** , **_ReturnAddressFromTop_** , **_DefaultFaceUp_** , **_DefaultOrientation_** , **_PrintEPostage_** , **_Vertical_** , **_RecipientNamefromLeft_** , **_RecipientNamefromTop_** , **_RecipientPostalfromLeft_** , **_RecipientPostalfromTop_** , **_SenderNamefromLeft_** , **_SenderNamefromTop_** , **_SenderPostalfromLeft_** , **_SenderPostalfromTop_** )

 _expression_ Required. A variable that represents an **[Envelope](envelope-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _ExtractAddress_|Optional| **Variant**| **True** to use the text marked by the "EnvelopeAddress" bookmark (a user-defined bookmark) as the recipient's address.|
| _Address_|Optional| **Variant**|A string that specifies the recipient's address (ignored if ExtractAddress is  **True** ).|
| _AutoText_|Optional| **Variant**|The name of the AutoText entry that includes a recipient's address.|
| _OmitReturnAddress_|Optional| **Variant**| **True** to omit the return address.|
| _ReturnAddress_|Optional| **Variant**|A string that specifies the return address.|
| _ReturnAutoText_|Optional| **Variant**|The name of the AutoText entry that includes a return address.|
| _PrintBarCode_|Optional| **Variant**| **True** to add a POSTNET bar code. For U.S. mail only.|
| _PrintFIMA_|Optional| **Variant**| **True** to add a Facing Identification Mark (FIM-A) for use in presorting courtesy reply mail. For U.S. mail only.|
| _Size_|Optional| **Variant**|A string that specifies the envelope size. The string should match one of the sizes listed on the left side of the Envelope size box in the  **Envelope Options** dialog box (for example, "Size 10").|
| _Height_|Optional| **Variant**|The height of the envelope (in points) when the Size argument is set to "Custom size".|
| _Width_|Optional| **Variant**|The width of the envelope (in points) when the Size argument is set to "Custom size".|
| _FeedSource_|Optional| **Variant**| **True** to use the **FeedSource** property of the **Envelope** object to specify which paper tray to use when printing the envelope.|
| _AddressFromLeft_|Optional| **Variant**|The distance (in points) between the left edge of the envelope and the recipient's address.|
| _AddressFromTop_|Optional| **Variant**|The distance (in points) between the top edge of the envelope and the recipient's address.|
| _ReturnAddressFromLeft_|Optional| **Variant**|The distance (in points) between the left edge of the envelope and the return address.|
| _ReturnAddressFromTop_|Optional| **Variant**|The distance (in points) between the top edge of the envelope and the return address.|
| _DefaultFaceUp_|Optional| **Variant**| **True** to print the envelope face up; **False** to print it face down.|
| _DefaultOrientation_|Optional| **Variant**|The orientation of the envelope. Can be any  **WdEnvelopeOrientation** constant.|
| _PrintEPostage_|Optional| **Variant**| **True** to print postage using an Internet e-postage vendor.|
| _Vertical_|Optional| **Variant**| **True** prints text vertically on the envelope. Used for Asian-language envelopes.|
| _RecipientNamefromLeft_|Optional| **Variant**|The position of the recipient's name, measured in points, from the left edge of the envelope. Used for Asian-language envelopes.|
| _RecipientNamefromTop_|Optional| **Variant**|The position of the recipient's name, measured in points, from the top edge of the envelope. Used for Asian-language envelopes.|
| _RecipientPostalfromLeft_|Optional| **Variant**|The position of the recipient's postal code, measured in points, from the left edge of the envelope. Used for Asian-language envelopes.|
| _RecipientPostalfromTop_|Optional| **Variant**|The position of the recipient's postal code, measured in points, from the top edge of the envelope. Used for Asian-language envelopes.|
| _SenderNamefromLeft_|Optional| **Variant**|The position of the sender's name, measured in points, from the left edge of the envelope. Used for Asian-language envelopes.|
| _SenderNamefromTop_|Optional| **Variant**|The position of the sender's name, measured in points, from the top edge of the envelope. Used for Asian-language envelopes.|
| _SenderPostalfromLeft_|Optional| **Variant**|The position of the sender's postal code, measured in points, from the left edge of the envelope. Used for Asian-language envelopes.|
| _SenderPostalfromTop_|Optional| **Variant**|The position of the sender's postal code, measured in points, from the top edge of the envelope. Used for Asian-language envelopes.|

## Example

This example prints an envelope using the user address as the return address and a predefined recipient address.


```
recep = "Don Funk" &; vbCr &; "123 Skye St." &; vbCr &; _ 
    "OurTown, WA 98107" 
ActiveDocument.Envelope.PrintOut Address:=recep, _ 
    ReturnAddress:=Application.UserAddress, _ 
    Size:="Size 10", PrintBarCode:=True
```


## See also


#### Concepts


[Envelope Object](envelope-object-word.md)

