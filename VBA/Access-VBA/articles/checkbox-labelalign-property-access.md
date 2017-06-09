---
title: CheckBox.LabelAlign Property (Access)
keywords: vbaac10.chm10729
f1_keywords:
- vbaac10.chm10729
ms.prod: access
api_name:
- Access.CheckBox.LabelAlign
ms.assetid: 255be436-51d3-0926-a7ce-a5b595ff59ce
ms.date: 06/08/2017
---


# CheckBox.LabelAlign Property (Access)

The property specifies the text alignment within attached labels on new controls. Read/write  **Byte**.


## Syntax

 _expression_. **LabelAlign**

 _expression_ A variable that represents a **CheckBox** object.


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


[CheckBox Object](checkbox-object-access.md)

