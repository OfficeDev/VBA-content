---
title: Application.BoxStylesEditEx Method (Project)
keywords: vbapj.chm2154
f1_keywords:
- vbapj.chm2154
ms.prod: project-server
api_name:
- Project.Application.BoxStylesEditEx
ms.assetid: 8a473e08-7893-6871-d015-23e1791e67e3
ms.date: 06/08/2017
---


# Application.BoxStylesEditEx Method (Project)

Sets the style of boxes in the Network Diagram view, where colors can be hexadecimal values.


## Syntax

 _expression_. **BoxStylesEditEx**( ** _Style_**, ** _DataTemplate_**, ** _HorizontalGridlines_**, ** _VerticalGridlines_**, ** _BorderShape_**, ** _BorderColor_**, ** _BorderWidth_**, ** _BackgroundColor_**, ** _BackgroundPattern_** )

 _expression_ An expression that returns an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Style_|Required|**Long**|The box style to change. Can be one of the  **[PjBoxStyle](pjboxstyle-enumeration-project.md)** constants.|
| _DataTemplate_|Optional|**String**|The name of the data template to use for the style.|
| _HorizontalGridlines_|Optional|**Boolean**|**True** if horizontal gridlines separate each row in the box; otherwise, **False**.|
| _VerticalGridlines_|Optional|**Boolean**|**True** if vertical gridlines separate each row in the box; otherwise, **False**.|
| _BorderShape_|Optional|**Long**|The shape of the box border. Can be one of the  **[PjBoxShape](pjboxshape-enumeration-project.md)** constants.|
| _BorderColor_|Optional|**Long**|The color of the box border. Can be a hexadecimal value for the RGB color, where red is the last byte. For example, the value &;HFF0000 is blue and &;H00FFFF is yellow. |
| _BorderWidth_|Optional|**Long**|A value from 1 through 4 that specifies the width of the box border, in pixels.|
| _BackgroundColor_|Optional|**Long**|The color of the box background. Can be a hexadecimal value for the RGB color.|
| _BackgroundPattern_|Optional|**Long**|The pattern for the background. Can be one of the [PjBackgroundPattern](pjbackgroundpattern-enumeration-project.md) constants.|

### Return Value

 **Boolean**


## Remarks

To display the  **Box Styles** dialog box, use the **[BarBoxStyles](application-barboxstyles-method-project.md)** method.


## Example

The following example changes boxes with the  **pjBoxCritical** style to be shown as rounded rectangles, adds vertical gridlines, sets the border color to a dark red, and sets the background color to light gray with a dither pattern.


```vb
Sub BoxStyles_EditCritical() 
 'Activate the Network Diagram view 
 ViewApply Name:="Network Diagram" 
 
 BoxStylesEditEx Style:=pjBoxCritical, BorderShape:=pjBoxRoundedRectangle, VerticalGridlines:=True, _ 
 BorderColor:=&;HB0, BorderWidth:=3, _ 
 BackgroundColor:=&;HE0E0E0, BackgroundPattern:=pjBackgroundLightDither 
End Sub
```


 **Note**  If you use any of the  **PjColor** enumeration constants for the _BorderColor_ or _BackgroundColor_ parameters, the color will be nearly black. For example, the value of **pjGreen** is 9, which in the **BoxLinksEx** method is a very dark red. To use only the sixteen colors available with **PjColor** constants, use the[BoxStylesEdit](application-boxstylesedit-method-project.md) method.


