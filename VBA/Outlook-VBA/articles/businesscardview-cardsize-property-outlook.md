---
title: BusinessCardView.CardSize Property (Outlook)
keywords: vbaol11.chm2937
f1_keywords:
- vbaol11.chm2937
ms.prod: outlook
api_name:
- Outlook.BusinessCardView.CardSize
ms.assetid: 0a1cbe6d-cc1a-1701-fe43-8704002b2212
ms.date: 06/08/2017
---


# BusinessCardView.CardSize Property (Outlook)

Returns or sets a  **Long** value that represents the size, as a percentage, of an Electronic Business Card (EBC) in the view. Read/write.


## Syntax

 _expression_ . **CardSize**

 _expression_ An expression that returns a **BusinessCardView** object.


## Remarks

This property can be set to a value between 20 and 100. If this property is set to a value less than 20, the property is set to 20. If this property is set to a value greater than 100, the property is set to 100.


## See also


#### Concepts


[BusinessCardView Object](businesscardview-object-outlook.md)

