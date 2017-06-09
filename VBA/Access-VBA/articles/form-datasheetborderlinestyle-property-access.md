---
title: Form.DatasheetBorderLineStyle Property (Access)
keywords: vbaac10.chm13512
f1_keywords:
- vbaac10.chm13512
ms.prod: access
api_name:
- Access.Form.DatasheetBorderLineStyle
ms.assetid: 8a752955-97fe-933a-4130-62f63dbf6566
ms.date: 06/08/2017
---


# Form.DatasheetBorderLineStyle Property (Access)

Returns or sets a  **Byte** indicating the line style to use for the border of the specified datasheet. Read/write.


## Syntax

 _expression_. **DatasheetBorderLineStyle**

 _expression_ A variable that represents a **Form** object.


## Remarks

Valid values are between zero and eight. Values greater than eight are ignored; negative values or values above 255 cause an error.



|**Value**|**Description**|
|:-----|:-----|
|0|Transparent border|
|1|Solid|
|2|Dashes|
|3|Short dashes|
|4|Dots|
|5|Sparse dots|
|6|Dash-dot|
|7|Dash-dot-dot|
|8|Double solid|

## Example

This example sets the datasheet border line style on the first open form to short dashes. The form must be set to Datasheet View in order for you to see the change.


```vb
Forms(0).DatasheetBorderLineStyle = 3 

```


## See also


#### Concepts


[Form Object](form-object-access.md)

