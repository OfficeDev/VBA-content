---
title: Application.BoxStylesEdit Method (Project)
keywords: vbapj.chm2387
f1_keywords:
- vbapj.chm2387
ms.prod: project-server
api_name:
- Project.Application.BoxStylesEdit
ms.assetid: 21a15566-3ee2-521a-f813-0f0baa806bfd
ms.date: 06/08/2017
---


# Application.BoxStylesEdit Method (Project)

Sets the style of boxes in the Network Diagram view.


## Syntax

 _expression_. **BoxStylesEdit**( ** _Style_**, ** _DataTemplate_**, ** _HorizontalGridlines_**, ** _VerticalGridlines_**, ** _BorderShape_**, ** _BorderColor_**, ** _BorderWidth_**, ** _BackgroundColor_**, ** _BackgroundPattern_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Style_|Required|**Long**|The box style to change. Can be one of the  **[PjBoxStyle](pjboxstyle-enumeration-project.md)** constants.|
| _DataTemplate_|Optional|**String**|The name of the data template to use for the style.|
| _HorizontalGridlines_|Optional|**Boolean**|**True** if horizontal gridlines separate each row in the box; otherwise, **False**.|
| _VerticalGridlines_|Optional|**Boolean**|**True** if vertical gridlines separate each row in the box; otherwise, **False**.|
| _BorderShape_|Optional|**Long**|The shape of the box border. Can be one of the  **[PjBoxShape](pjboxshape-enumeration-project.md)** constants.|
| _BorderColor_|Optional|**Long**|The color of the box border. Can be one of the  **[PjColor](pjcolor-enumeration-project.md)** constants.|
| _BorderWidth_|Optional|**Long**|A value from 1 through 4 that specifies the width of the box border, in pixels.|
| _BackgroundColor_|Optional|**Long**|The color of the box background. Can be one of the  **[PjColor](pjcolor-enumeration-project.md)** constants.|
| _BackgroundPattern_|Optional|**Long**|The pattern for the background. Can be one of the  **[PjBackgroundPattern](pjbackgroundpattern-enumeration-project.md)** constants.|

### Return Value

 **Boolean**


## Remarks

To display the  **Box Styles** dialog box, use the **[BarBoxStyles](application-barboxstyles-method-project.md)** method.

To edit box link lines where the colors can be RGB values, use the  **[BoxStylesEditEx](application-boxstyleseditex-method-project.md)** method.


## Example

The following example changes boxes with the  **pjBoxCritical** style to be shown as rounded rectangles, adds vertical gridlines, and sets the border and background colors.


```vb
Sub BoxStyles_Edit() 
 'Activate the Network Diagram view 
 ViewApply Name:="Network Diagram" 
 
 BoxStylesEdit Style:=pjBoxCritical, BorderShape:=pjBoxRoundedRectangle, VerticalGridlines:=True, _ 
 BorderColor:=pjRed, BorderWidth:=3, _ 
 BackgroundColor:=pjGray, BackgroundPattern:=pjBackgroundLightDither 
End Sub
```


