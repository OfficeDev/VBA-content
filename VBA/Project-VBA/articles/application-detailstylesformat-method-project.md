---
title: Application.DetailStylesFormat Method (Project)
keywords: vbapj.chm962
f1_keywords:
- vbapj.chm962
ms.prod: project-server
api_name:
- Project.Application.DetailStylesFormat
ms.assetid: df3b7963-134f-be55-715e-2e4c214b35fc
ms.date: 06/08/2017
---


# Application.DetailStylesFormat Method (Project)

Sets the format of timescaled data fields in a Resource Usage view or Task Usage view.


## Syntax

 _expression_. **DetailStylesFormat**( ** _Item_**, ** _Font_**, ** _Size_**, ** _Bold_**, ** _Italic_**, ** _Underline_**, ** _Color_**, ** _CellColor_**, ** _Pattern_**, ** _ShowInMenu_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Item_|Optional|**Long**|The timescaled data field to format. If the active view is the Task Usage view, the value can be one of the  **[PjTaskTimescaledData](pjtasktimescaleddata-enumeration-project.md)** constants. If the active view is the Resource Usage view, the value can be one of the **[PjResourceTimescaledData](pjresourcetimescaleddata-enumeration-project.md)** constants.|
| _Font_|Optional|**String**|The name of the font.|
| _Size_|Optional|**Integer**|The size of the font in points.|
| _Bold_|Optional|**Boolean**|**True** if the font is bold; otherwise, **False**.|
| _Italic_|Optional|**Boolean**|**True** if the font is italic; otherwise, **False**.|
| _Underline_|Optional|**Boolean**|**True** if the font is underlined; otherwise, **False**.|
| _Color_|Optional|**Long**|The color of the font. Can be one of the  **[PjColor](pjcolor-enumeration-project.md)** constants.|
| _CellColor_|Optional|**Long**|The color of the cell background. Can be one of the  **PjColor** constants.|
| _Pattern_|Optional|**Long**|The pattern for nonworking times. Can be one of the  **[PjFillPattern](pjfillpattern-enumeration-project.md)** constants.|
| _ShowInMenu_|Optional|**Boolean**|**True** if the field specified with **Item** appears in the shortcut menu; otherwise, **False**. The default value is **False**.|

### Return Value

 **Boolean**


## Remarks

Using the  **DetailStylesFormat** method without specifying any arguments displays the **Detail Styles** dialog box with the **Usage Details** tab selected.

To edit the timescale data where the text and cell color can be a hexadecimal RGB value, and the font can include the strikethrough style, use the  **[DetailStylesFormatEx](application-detailstylesformatex-method-project.md)** method.


## Example

The following example makes overallocations stand out from other information in a usage view.


```vb
Sub HighlightOverallocations() 
 DetailStylesAdd pjOverallocation 
 DetailStylesFormat Item:=pjOverallocation, Font:="Arial", Size:=10, _ 
 Bold:=True, Color:=pjRed, CellColor:=pjBlack, Pattern:=pjSolidFill 
End Sub
```


