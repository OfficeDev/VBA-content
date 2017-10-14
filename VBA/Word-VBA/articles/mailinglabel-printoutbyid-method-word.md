---
title: MailingLabel.PrintOutByID Method (Word)
keywords: vbawd10.chm152502379
f1_keywords:
- vbawd10.chm152502379
ms.prod: word
api_name:
- Word.MailingLabel.PrintOutByID
ms.assetid: 841a5c10-e6e7-b852-a947-e7e450537a9e
ms.date: 06/08/2017
---


# MailingLabel.PrintOutByID Method (Word)

Prints a label or a page of labels with the same address.


## Syntax

 _expression_ . **PrintOutByID**( **_LabelID_** , **_Address_** , **_ExtractAddress_** , **_LaserTray_** , **_SingleLabel_** , **_Row_** , **_Column_** , **_PrintEPostageLabel_** , **_Vertical_** )

 _expression_ An expression that returns a **[MailingLabel](mailinglabel-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _LabelID_|Optional| **Variant**|The mailing label identification.|
| _Address_|Optional| **Variant**|The text for the label address.|
| _ExtractAddress_|Optional| **Variant**| **True** to use the text marked by the "EnvelopeAddress" bookmark (a user-defined bookmark) as the label text. If this argument is specified, Address and AutoText are ignored.|
| _LaserTray_|Optional| **Variant**|The laser printer tray to be used. Can be any  **WdPaperTray** constant.|
| _SingleLabel_|Optional| **Variant**| **True** to print a single label; **False** to print an entire page of the same label.|
| _Row_|Optional| **Variant**|The label row for a single label. Not valid if SingleLabel is  **False** .|
| _Column_|Optional| **Variant**|The label column for a single label. Not valid if SingleLabel is  **False** .|
| _PrintEPostageLabel_|Optional| **Variant**| **True** to print postage using an Internet e-postage vendor.|
| _Vertical_|Optional| **Variant**| **True** prints text vertically on the label. Used for Asian-language mailing labels.|

## See also


#### Concepts


[MailingLabel Object](mailinglabel-object-word.md)

