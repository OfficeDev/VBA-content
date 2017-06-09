---
title: Application.FontEx Method (Project)
keywords: vbapj.chm937
f1_keywords:
- vbapj.chm937
ms.prod: project-server
api_name:
- Project.Application.FontEx
ms.assetid: 4904d4b1-dacb-8020-0c4e-3af0503c68ba
ms.date: 06/08/2017
---


# Application.FontEx Method (Project)

Sets the font for the text in the active cells.


## Syntax

 _expression_. **FontEx**( ** _Name_**, ** _Size_**, ** _Bold_**, ** _Italic_**, ** _Underline_**, ** _Color_**, ** _Reset_**, ** _CellColor_**, ** _Pattern_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Name_|Optional|**String**|The name of the font.|
| _Size_|Optional|**Integer**|The size of the font in points.|
| _Bold_|Optional|**Variant**|**True** if the font is bold.|
| _Italic_|Optional|**Boolean**|**True** if the font is italic; otherwise, **False**.|
| _Underline_|Optional|**Boolean**|**True** if the font is underlined; otherwise, **False**.|
| _Color_|Optional|**Long**|The color of the font. Can be one of the  **[PjColor](pjcolor-enumeration-project.md)** constants.|
| _Reset_|Optional|**Boolean**|**True** if the font is reset to its default characteristics. All other arguments are ignored. The default value is **False**.|
| _CellColor_|Optional|**Variant**|The color of the cell. Can be one of the  **[PjColor](pjcolor-enumeration-project.md)** constants.|
| _Pattern_|Optional|**Variant**|Background pattern of the cell. Can be one of the  **[PjFillPattern](pjfillpattern-enumeration-project.md)** constants.|

### Return Value

 **Boolean**


## Remarks

To set the font with an RGB hexadecimal value for color or with a strikethrough format, use the  **[Font32Ex](application-font32ex-method-project.md)** method.


## Example

The following example formats selected text using 16-point Tahoma in a red color.


```vb
Sub FormatTahoma16() 
 FontEx Name:="Tahoma", Size:=16, Color:=pjRed 
End Sub
```


