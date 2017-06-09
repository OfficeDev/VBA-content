---
title: ComboBox.DropButtonStyle Property (Outlook Forms Script)
keywords: olfm10.chm2001110
f1_keywords:
- olfm10.chm2001110
ms.prod: outlook
ms.assetid: 91cf54d6-1378-8cf5-6a2c-153d2ef4221e
ms.date: 06/08/2017
---


# ComboBox.DropButtonStyle Property (Outlook Forms Script)

Returns or sets a  **fmDropButtonStyle** value that represents the symbol displayed on the drop button in a **[ComboBox](combobox-object-outlook-forms-script.md)**. Read/write.


## Syntax

 _expression_. **DropButtonStyle**

 _expression_A variable that represents a  **ComboBox** object.


## Remarks

The possible values for  **DropButtonStyle** are:



|**Value**|**Description**|
|:-----|:-----|
|0|Displays a plain button, with no symbol.|
|1|Displays a down arrow (default).|
|2|Displays an ellipsis (...).|
|3|Displays a horizontal line like an underscore character.|
The recommended setting for showing items in a list is 1.


