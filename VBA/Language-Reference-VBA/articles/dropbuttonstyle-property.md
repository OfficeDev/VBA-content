---
title: DropButtonStyle Property
keywords: fm20.chm2001110
f1_keywords:
- fm20.chm2001110
ms.prod: office
api_name:
- Office.DropButtonStyle
ms.assetid: 14d5061f-1267-64b5-5734-65febe6e584c
ms.date: 06/08/2017
---


# DropButtonStyle Property



Specifies the symbol displayed on the drop button in a  **ComboBox**.
 **Syntax**
 _object_. **DropButtonStyle** [= _fmDropButtonStyle_ ]
The  **DropButtonStyle** property syntax has these parts:


|**Part**|**Description**|
|:-----|:-----|
| _object_|Required. A valid object.|
| _fmDropButtonStyle_|Optional. The appearance of the drop button.|
 **Settings**
The settings for  _fmDropButtonStyle_ are:


|**Constant**|**Value**|**Description**|
|:-----|:-----|:-----|
| _fmDropButtonStylePlain_|0|Displays a plain button, with no symbol.|
| _fmDropButtonStyleArrow_|1|Displays a down arrow (default).|
| _fmDropButtonStyleEllipsis_|2|Displays an ellipsis (...).|
| _fmDropButtonStyleReduce_|3|Displays a horizontal line like an underscore character.|
 **Remarks**
The recommended setting for showing items in a list is  **fmDropButtonStyleArrow**. If you want to use the drop button in another way, such as to display a dialog box, specify **fmDropButtonStyleEllipsis**, **fmDropButtonStylePlain**, or **fmDropButtonStyleReduce** and trap the DropButtonClick event.

