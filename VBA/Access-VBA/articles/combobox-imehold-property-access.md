---
title: ComboBox.IMEHold Property (Access)
keywords: vbaac10.chm11390
f1_keywords:
- vbaac10.chm11390
ms.prod: access
api_name:
- Access.ComboBox.IMEHold
ms.assetid: ab128652-1de6-e4a2-4bc5-99936b3fee7f
ms.date: 06/08/2017
---


# ComboBox.IMEHold Property (Access)

[Language-specific information](learn-about-language-specific-information-access.md)You can use the  **IMEHold/Hold KanjiConversionMode** property to show whether the Kanji Conversion Mode is maintained when the control loses the focus. Read/write **Boolean**.


## Syntax

 _expression_. **IMEHold**

 _expression_ A variable that represents a **ComboBox** object.


## Remarks

The  **IMEHold/Hold KanjiConversionMode** property uses the following settings.



|**Setting**|**Description**|
|:-----|:-----|
|True|Maintains the Kanji Conversion Mode set in the last control.|
|False|Does not maintain the Kanji Conversion Mode set in the last control (default).|
By setting the  **IMEMode/KanjiConversionMode** property, you can designate whether or not the Kanji Conversion Mode is maintained when the control loses the focus. If this property is set to Yes, the Kanji Conversion Mode setting is maintained when the control loses the focus and once that control regains the focus, the Kanji Conversion Mode setting for that control is restored. If this property is set to No (default setting), the Kanji Conversion Mode will be set by the **IMEMode/KanjiConversionMode** property for that control.


 **Note**  To set the Kanji Conversion Mode when the focus shifts to the control, set the  **IMEMode/KanjiConversionMode** property.


## See also


#### Concepts


[ComboBox Object](combobox-object-access.md)

