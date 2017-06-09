---
title: MailingLabel.CreateNewDocument Method (Word)
keywords: vbawd10.chm152502376
f1_keywords:
- vbawd10.chm152502376
ms.prod: word
api_name:
- Word.MailingLabel.CreateNewDocument
ms.assetid: a52831c0-42cb-e970-14e7-2c71bcc5c2f1
ms.date: 06/08/2017
---


# MailingLabel.CreateNewDocument Method (Word)

Creates a new label document using either the default label options or ones that you specify. Returns a  **Document** object that represents the new document.


## Syntax

 _expression_ . **CreateNewDocument**( **_Name_** , **_Address_** , **_AutoText_** , **_ExtractAddress_** , **_LaserTray_** , **_PrintEPostageLabel_** , **_Vertical_** )

 _expression_ Required. A variable that represents a **[MailingLabel](mailinglabel-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Name_|Optional| **Variant**|The mailing label name.|
| _Address_|Optional| **Variant**|The text for the mailing label.|
| _AutoText_|Optional| **Variant**|The name of the AutoText entry that includes the mailing label text.|
| _ExtractAddress_|Optional| **Variant**| **True** to use the address text marked by the user-defined bookmark named "EnvelopeAddress" instead of using the Address argument.|
| _LaserTray_|Optional| **Variant**|The laser printer tray. Can be one of the  **[WdPaperTray](wdpapertray-enumeration-word.md)** constants.|
| _PrintEPostageLabel_|Optional| **Variant**| **True** to print postage using an Internet e-postage vendor.|
| _Vertical_|Optional| **Variant**| **True** formats text vertically on the label. Used for Asian-language mailing labels.|

### Return Value

Document


## Example

This example creates a new Avery 2160 minilabel document using a predefined address.


```
addr = "Dave Edson" &; vbCr &; "123 Skye St." _ 
 &; vbCr &; "Our Town, WA 98004" 
Application.MailingLabel.CreateNewDocument _ 
 Name:="2160 mini", Address:=addr, ExtractAddress:=False
```

This example creates a new Avery 5664 shipping-label document using the selected text as the address.




```
addr = Selection.Text 
Application.MailingLabel.CreateNewDocument _ 
 Name:="5664", Address:=addr, _ 
 LaserTray:=wdPrinterUpperBin
```

This example creates a new self-adhesive-label document using the EnvelopeAddress bookmark text as the address.




```vb
If ActiveDocument.Bookmarks.Exists("EnvelopeAddress") = True Then 
 Application.MailingLabel.CreateNewDocument _ 
 Name:="Self Adhesive Tab 1 1/2""", ExtractAddress:=True 
End If
```


## See also


#### Concepts


[MailingLabel Object](mailinglabel-object-word.md)

