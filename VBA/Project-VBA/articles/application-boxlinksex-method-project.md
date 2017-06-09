---
title: Application.BoxLinksEx Method (Project)
keywords: vbapj.chm2158
f1_keywords:
- vbapj.chm2158
ms.prod: project-server
api_name:
- Project.Application.BoxLinksEx
ms.assetid: f6292e01-3f4a-3b83-e86c-2316c83b2509
ms.date: 06/08/2017
---


# Application.BoxLinksEx Method (Project)

Specifies the appearance of link lines in the active Network Diagram view, where colors can be hexadecimal values.


## Syntax

 _expression_. **BoxLinksEx**( ** _Style_**, ** _ShowArrows_**, ** _ShowLabels_**, ** _ColorMode_**, ** _CriticalColor_**, ** _NoncriticalColor_** )

 _expression_ An expression that returns an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Style_|Optional|**Long**|Specifies the style of link lines. Can be one of the following  **[PjLinkStyle](pjlinkstyle-enumeration-project.md)** constants: **pjLinkStraight** or **pjLinkRectilinear**.|
| _ShowArrows_|Optional|**Boolean**|**True** if link lines have arrows showing the direction of the link; otherwise, **False**.|
| _ShowLabels_|Optional|**Boolean**|**True** if link lines have labels showing the link type (FS, SS, SF, or FF); otherwise, **False**.|
| _ColorMode_|Optional|**Long**|Specifies how the color of link lines is determined. Can be one of the  **[PjLinkColorMode](pjlinkcolormode-enumeration-project.md)** constants.|
| _CriticalColor_|Optional|**Long**|The color of link lines between critical tasks. Can be a hexadecimal value for the RGB color, where red is the last byte. For example, the value &;HFF0000 is blue and &;H00FFFF is yellow. The default value is 0, which does not change the previous color.|
| _NoncriticalColor_|Optional|**Long**| The color of link lines between noncritical tasks. Can be a hexadecimal value; the default value is 0, which does not change the previous color.|

### Return Value

 **Boolean**


## Remarks

If no arguments are specified, the  **BoxLinksEx** method has no effect. If _ColorMode_ is **pjColorModePredecessor**, the _NoncriticalColor_ and _CriticalColor_ parameters are ignored.


## Example

The following example shows link labels and then sets critical links to a purple color and noncritical links to a teal color.


```vb
Sub BoxLink_ChangeColor() 
    'Activate the Network Diagram view 
    ViewApply Name:="Network Diagram" 
 
    BoxLinksEx Style:=pjLinkRectilinear, ShowArrows:=True, ShowLabels:=True, ColorMode:=pjColorModeCustom, _ 
        CriticalColor:=&;HBB00BB, noncriticalcolor:=&;H999900 
End Sub
```


 **Note**  If you use any of the  **PjColor** enumeration constants for the _CriticalColor_ or _NoncriticalColor_ parameters, the color will be nearly black. For example, the value of **pjGreen** is 9, which in the **BoxLinksEx** method is a very dark red. To use only the sixteen colors available with **PjColor** constants, use the[BoxLinks](application-boxlinks-method-project.md) method.


