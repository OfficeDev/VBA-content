---
title: NavigationButton.LabelAlign Property (Access)
keywords: vbaac10.chm10486
f1_keywords:
- vbaac10.chm10486
ms.prod: access
api_name:
- Access.NavigationButton.LabelAlign
ms.assetid: d6562f66-5b9a-1f91-e140-b84a57ea5ff9
ms.date: 06/08/2017
---


# NavigationButton.LabelAlign Property (Access)

The property specifies the text alignment within attached labels on new controls. Read/write  **Byte**.


## Syntax

 _expression_. **LabelAlign**

 _expression_ A variable that represents a **NavigationButton** object.


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


[NavigationButton Object](navigationbutton-object-access.md)

