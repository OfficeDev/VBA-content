---
title: OptionButton.LabelAlign Property (Access)
keywords: vbaac10.chm10602
f1_keywords:
- vbaac10.chm10602
ms.prod: access
api_name:
- Access.OptionButton.LabelAlign
ms.assetid: 3ecb4d1f-7e32-9699-b2c3-6918d7b2eb61
ms.date: 06/08/2017
---


# OptionButton.LabelAlign Property (Access)

The property specifies the text alignment within attached labels on new controls. Read/write  **Byte**.


## Syntax

 _expression_. **LabelAlign**

 _expression_ A variable that represents an **OptionButton** object.


## Remarks

The  **LabelAlign** property uses the following settings.



|**Setting**|**Description**|
|:-----|:-----|
|0|(Default) The label text aligns to the left.|
|1|The label text aligns to the left.|
|2|The label text is centered.|
|3|The label text aligns to the right.|
|4|The label text is evenly distributed.|
You can set the  **LabelAlign** property by using a control's default control style or the **DefaultControl** property in Visual Basic.

When created, controls have an attached label (as long as their  **AutoLabel** property is set to Yes). Changes to the **LabelAlign** default control style setting affect only controls created on the current form or report.


## See also


#### Concepts


[OptionButton Object](optionbutton-object-access.md)

