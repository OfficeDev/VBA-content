---
title: CardView.MultiLineFieldHeight Property (Outlook)
keywords: vbaol11.chm2601
f1_keywords:
- vbaol11.chm2601
ms.prod: outlook
api_name:
- Outlook.CardView.MultiLineFieldHeight
ms.assetid: 71b87b15-ef48-9214-295c-731bb9fbc808
ms.date: 06/08/2017
---


# CardView.MultiLineFieldHeight Property (Outlook)

Returns or sets a  **Long** value that determines the minimum number of lines for multiline fields displayed in the **[CardView](cardview-object-outlook.md)** object. Read/write.


## Syntax

 _expression_ . **MultiLineFieldHeight**

 _expression_ A variable that represents a **CardView** object.


## Remarks

This property can be set to a value between 1 and 20. If this property is set to a value less than 1, the property is set to 1. If this property is set to a value greater than 20, the property is set to 20. The default value for this property is 1.


## See also


#### Concepts


[CardView Object](cardview-object-outlook.md)

