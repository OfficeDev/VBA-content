---
title: ToggleButton.LabelAlign Property (Access)
keywords: vbaac10.chm11740
f1_keywords:
- vbaac10.chm11740
ms.prod: access
api_name:
- Access.ToggleButton.LabelAlign
ms.assetid: fa8b44e8-9e42-8088-e369-a176bb320a05
ms.date: 06/08/2017
---


# ToggleButton.LabelAlign Property (Access)

The property specifies the text alignment within attached labels on new controls. Read/write  **Byte**.


## Syntax

 _expression_. **LabelAlign**

 _expression_ A variable that represents a **ToggleButton** object.


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


[ToggleButton Object](togglebutton-object-access.md)

