---
title: Application.DetailStylesFormatEx Method (Project)
keywords: vbapj.chm2164
f1_keywords:
- vbapj.chm2164
ms.prod: project-server
api_name:
- Project.Application.DetailStylesFormatEx
ms.assetid: 3e460e76-ff7b-f07b-058c-1e37c53e453e
ms.date: 06/08/2017
---


# Application.DetailStylesFormatEx Method (Project)

Sets the format of timescaled data fields in a Resource Usage view or Task Usage view, where colors can be hexadecimal values.


## Syntax

 _expression_. **DetailStylesFormatEx**( ** _Item_**, ** _Font_**, ** _Size_**, ** _Bold_**, ** _Italic_**, ** _Underline_**, ** _Color_**, ** _CellColor_**, ** _Pattern_**, ** _ShowInMenu_**, ** _Strikethrough_** )

 _expression_ An expression that returns an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Item_|Optional|**Long**|The timescaled data field to format. If the active view is the Task Usage view, the value can be one of the  **[PjTaskTimescaledData](pjtasktimescaleddata-enumeration-project.md)** constants. If the active view is the Resource Usage view, the value can be one of the **[PjResourceTimescaledData](pjresourcetimescaleddata-enumeration-project.md)** constants.|
| _Font_|Optional|**String**|The name of the font.|
| _Size_|Optional|**Integer**|The size of the font in points.|
| _Bold_|Optional|**Boolean**|**True** if the font is bold; otherwise, **False**.|
| _Italic_|Optional|**Boolean**|**True** if the font is italic; otherwise, **False**.|
| _Underline_|Optional|**Boolean**|**True** if the font is underlined; otherwise, **False**.|
| _Color_|Optional|**Long**|The color of the font. Can be a hexadecimal value, where red is the last byte. For example, the value &;HFF0000 is blue and &;H00FFFF is yellow.|
| _CellColor_|Optional|**Long**|The color of the cell background. Can be a hexadecimal value, where red is the last byte. For example, the value &;HFF00 is green.|
| _Pattern_|Optional|**Long**|The pattern for nonworking times. Can be one of the  **[PjFillPattern](pjfillpattern-enumeration-project.md)** constants.|
| _ShowInMenu_|Optional|**Boolean**|**True** if the field specified with **Item** appears in the shortcut menu; otherwise, **False**. The default value is **False**.|
| _Strikethrough_|Optional|**Variant**|**True** if the font is the strikethrough style.|

### Return Value

 **Boolean**


## Remarks

Using the  **DetailStylesFormat** method without specifying any arguments displays the **Detail Styles** dialog box with the **Usage Details** tab selected.


## Example

The following example makes overallocations stand out from other information in a usage view.


```vb
Sub HighlightOverallocations() 
    DetailStylesAdd pjOverallocation 
    DetailStylesFormatEx Item:=pjOverallocation, Font:="Arial", Size:=10, _ 
        Bold:=True, Color:=&;HA0, CellColor:=&;HFFB0B0, Pattern:=pjSolidFill 
End Sub
```


 **Note**  If you use any of the  **PjColor** enumeration constants for the _Color_ or _CellColor_ parameters, the color will be nearly black. For example, the value of **pjGreen** is 9, which in the **DetailStylesFormatEx** method is a very dark red. To use only the sixteen colors available with **PjColor** constants, use the **[DetailStylesFormat](application-detailstylesformat-method-project.md)** method.


