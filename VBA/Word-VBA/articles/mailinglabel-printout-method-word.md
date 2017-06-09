---
title: MailingLabel.PrintOut Method (Word)
keywords: vbawd10.chm152502377
f1_keywords:
- vbawd10.chm152502377
ms.prod: word
api_name:
- Word.MailingLabel.PrintOut
ms.assetid: 3519226b-1c5f-8343-62b1-7e275793ca3c
ms.date: 06/08/2017
---


# MailingLabel.PrintOut Method (Word)

Prints a label or a page of labels with the same address.


## Syntax

 _expression_ . **PrintOut**( **_Name_** , **_Address_** , **_ExtractAddress_** , **_LaserTray_** , **_SingleLabel_** , **_Row_** , **_Column_** , **_PrintEPostageLabel_** , **_Vertical_** )

 _expression_ Required. A variable that represents a **[MailingLabel](mailinglabel-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Name_|Optional| **Variant**|The mailing label name.|
| _Address_|Optional| **Variant**|The text for the label address.|
| _ExtractAddress_|Optional| **Variant**| **True** to use the text marked by the "EnvelopeAddress" bookmark (a user-defined bookmark) as the label text. If this argument is specified, Address and AutoText are ignored.|
| _LaserTray_|Optional| **Variant**|The laser printer tray to be used. Can be any  **WdPaperTray** constant.|
| _SingleLabel_|Optional| **Variant**| **True** to print a single label; **False** to print an entire page of the same label.|
| _Row_|Optional| **Variant**|The label row for a single label. Not valid if SingleLabel is  **False** .|
| _Column_|Optional| **Variant**|The label column for a single label. Not valid if SingleLabel is  **False** .|
| _PrintEPostageLabel_|Optional| **Variant**| **True** to print postage using an Internet e-postage vendor.|
| _Vertical_|Optional| **Variant**| **True** prints text vertically on the label. Used for Asian-language mailing labels.|

## Example

This example prints a page of Avery 5664 mailing labels, using the specified address.


```
addr = "Jane Doe" &; vbCr &; "123 Skye St." _ 
 &; vbCr &; "OurTown, WA 98107" 
Application.MailingLabel.PrintOut Name:="5664", Address:=addr
```


## See also


#### Concepts


[MailingLabel Object](mailinglabel-object-word.md)

