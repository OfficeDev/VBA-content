---
title: Application.Font32Ex Method (Project)
keywords: vbapj.chm2149
f1_keywords:
- vbapj.chm2149
ms.prod: project-server
api_name:
- Project.Application.Font32Ex
ms.assetid: 5f4928a6-d7b3-ff30-48ef-a5037dbeff21
ms.date: 06/08/2017
---


# Application.Font32Ex Method (Project)

Sets the font for the text in the active cells, where the text color can be a hexadecimal value.


## Syntax

 _expression_. **Font32Ex**( ** _Name_**, ** _Size_**, ** _Bold_**, ** _Italic_**, ** _Underline_**, ** _Color_**, ** _Reset_**, ** _CellColor_**, ** _Pattern_**, ** _Strikethrough_** )

 _expression_ An expression that returns an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Name_|Optional|**String**|The name of the font.|
| _Size_|Optional|**Integer**|The size of the font in points.|
| _Bold_|Optional|**Variant**|**True** if the font is bold.|
| _Italic_|Optional|**Boolean**|**True** if the font is italic; otherwise, **False**.|
| _Underline_|Optional|**Boolean**|**True** if the font is underlined; otherwise, **False**.|
| _Color_|Optional|**Long**|The color of the font. Can be a hexadecimal RGB value, where red is the last byte. For example, &;H0000FF is red.|
| _Reset_|Optional|**Boolean**|**True** if the font is reset to its default characteristics. All other arguments are ignored. The default value is **False**.|
| _CellColor_|Optional|**Variant**|The color of the cell. Can be a hexadecimal RGB value, where red is the last byte. For example, &;HFFFF99 is cyan.|
| _Pattern_|Optional|**Variant**|Background pattern of the cell. Can be one of the  **[PjFillPattern](pjfillpattern-enumeration-project.md)** constants.|
| _Strikethrough_|Optional|**Variant**|**True** if the font is the strikethrough format.|

### Return Value

 **Boolean**


## Remarks

For the Color and CellColor parameters, the decimal value -16777216 sets the color to automatic (black for text and white for the cell color). 


## Example

The following example formats text in the selected cells using 16-point Tahoma in a pink color, and sets the cell color to a light yellow.


```vb
Sub FormatTahoma16() 
    Font32Ex Name:="Tahoma", Size:=16, Color:=&;HFF88FF, CellColor:=&;H99FFFF 
End Sub
```


 **Note**  If you use any of the  **PjColor** constants for the Color or CellColor parameters, the color will be nearly black. For example, the value of **pjGreen** is 9, which in the **Fon32Ex** method is a very dark red. To use only the sixteen colors available with **PjColor** constants, use the **[FontEx](application-fontex-method-project.md)** method.


