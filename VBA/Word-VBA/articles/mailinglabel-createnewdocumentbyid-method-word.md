---
title: MailingLabel.CreateNewDocumentByID Method (Word)
keywords: vbawd10.chm152502378
f1_keywords:
- vbawd10.chm152502378
ms.prod: word
api_name:
- Word.MailingLabel.CreateNewDocumentByID
ms.assetid: 5b2d0b50-89cd-e37b-48d3-f4475009ba79
ms.date: 06/08/2017
---


# MailingLabel.CreateNewDocumentByID Method (Word)

Creates a new label document using either the default label options or ones that you specify. Returns a  **Document** object that represents the new document.


## Syntax

 _expression_ . **CreateNewDocumentByID**( **_LabelID_** , **_Address_** , **_AutoText_** , **_ExtractAddress_** , **_LaserTray_** , **_PrintEPostageLabel_** , **_Vertical_** )

 _expression_ An expression that returns a **[MailingLabel](mailinglabel-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _LabelID_|Optional| **Variant**|The mailing label identification.|
| _Address_|Optional| **Variant**|The text for the mailing label.|
| _AutoText_|Optional| **Variant**|The name of the AutoText entry that includes the mailing label text.|
| _ExtractAddress_|Optional| **Variant**| **True** to use the address text marked by the user-defined bookmark named "EnvelopeAddress" instead of using the Address argument.|
| _LaserTray_|Optional| **Variant**|The laser printer tray. Can be one of the  **[WdPaperTray](wdpapertray-enumeration-word.md)** constants.|
| _PrintEPostageLabel_|Optional| **Variant**| **True** to print postage using an Internet e-postage vendor.|
| _Vertical_|Optional| **Variant**| **True** formats text vertically on the label. Used for Asian-language mailing labels.|

### Return Value

Document


## See also


#### Concepts


[MailingLabel Object](mailinglabel-object-word.md)

