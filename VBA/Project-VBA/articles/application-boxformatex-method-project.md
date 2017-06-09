---
title: Application.BoxFormatEx Method (Project)
keywords: vbapj.chm2155
f1_keywords:
- vbapj.chm2155
ms.prod: project-server
api_name:
- Project.Application.BoxFormatEx
ms.assetid: 2cec4b32-3170-8d0b-f73e-5dc64e5ffa68
ms.date: 06/08/2017
---


# Application.BoxFormatEx Method (Project)

Formats individual boxes in the Network Diagram view (PERT chart), where colors can be specified with hexadecimal values.


## Syntax

 _expression_. **BoxFormatEx**( ** _ProjectName_**, ** _TaskID_**, ** _DataTemplate_**, ** _HorizontalGridlines_**, ** _VerticalGridlines_**, ** _BorderShape_**, ** _BorderColor_**, ** _BorderWidth_**, ** _BackgroundColor_**, ** _BackgroundPattern_**, ** _Reset_** )

 _expression_ An expression that returns an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _ProjectName_|Optional|**String**|The name of the project containing  **TaskID** when working with consolidated projects. The default value is the name of the active project.|
| _TaskID_|Optional|**Long**|The identification number of the task represented by the box to change. The default behavior is to change the boxes that represent one or more selected tasks.|
| _DataTemplate_|Optional|**String**|The name of the data template to use.|
| _HorizontalGridlines_|Optional|**Boolean**|**True** if horizontal gridlines separate each row in the box; otherwise, **False**.|
| _VerticalGridlines_|Optional|**Boolean**|**True** if vertical gridlines separate each column in the box; otherwise, **False**.|
| _BorderShape_|Optional|**Long**|The shape of the box border. Can be one of the  **[PjBoxShape](pjboxshape-enumeration-project.md)** constants.|
| _BorderColor_|Optional|**Long**|The color of the box border. Can be a hexadecimal value for the RGB color, where red is the last byte. For example, the value  `&;HFF0000` is blue and `&;H00FFFF` is yellow.|
| _BorderWidth_|Optional|**Long**|Specifies the box border width, where values can be 1 to 4 for the four line widths shown in the  **Format Box** dialog box.|
| _BackgroundColor_|Optional|**Long**|The color of the box background. Can be a hexadecimal value, where red is the last byte. For example, the value  `&;HFFFF00` is blue-green and `&;HFF00FF` is purple.|
| _BackgroundPattern_|Optional|**Long**|The pattern for the background. Can be one of the [PjBackgroundPattern](pjbackgroundpattern-enumeration-project.md) constants.|
| _Reset_|Optional|**Boolean**|**True** if the box formatting is reset to the default style as shown in the **Box Styles** dialog box. If **Reset** is **True**, all arguments except **ProjectName** and **TaskID** are ignored.|

### Return Value

 **Boolean**


## Remarks

If  **TaskID** is specified, the associated task cannot be hidden due to application of a filter or a collapsed outline structure.

Using the  **BoxFormatEx** method with no arguments displays the **Format Box** dialog box for the selected tasks. If no tasks are selected, the **BoxFormatEx** method has no effect.

Use the  **BoxFormatEx** method to change the formatting of boxes from their default styles. To define the default styles, use the **BoxStylesEdit** or the **BoxStylesEditEx** method.


## Example

The following example changes the box border color to red and the background color to a light blue dithered pattern.


```vb
Sub BoxFormat_Color() 
    'Activate the Network Diagram view
    ViewApply Name:="Network Diagram"

    BoxFormatEx TaskID:="2", bordershape:=pjBoxRoundedRectangle, VerticalGridlines:=False, _
        BorderWidth:=3, backgroundpattern:=pjBackgroundLightDither, _
        BackgroundColor:=&;HFF0000, BorderColor:=&;HFF
End Sub
```


 **Note**  If you use any of the  **PjColor** constants for the _BorderColor_ or _BackgroundColor_ parameters, the color will be nearly black. For example, the value of **pjGreen** is 9, which in the **BoxFormatEx** method is a very dark red. To use only the sixteen colors available with **PjColor** constants, use the[BoxFormat](application-boxformat-method-project.md) method.


