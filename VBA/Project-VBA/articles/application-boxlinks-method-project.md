---
title: Application.BoxLinks Method (Project)
keywords: vbapj.chm44
f1_keywords:
- vbapj.chm44
ms.prod: project-server
api_name:
- Project.Application.BoxLinks
ms.assetid: da12c972-9647-9e1f-2909-1e0a18aff32b
ms.date: 06/08/2017
---


# Application.BoxLinks Method (Project)

Specifies the appearance of link lines in the active Network Diagram view.


## Syntax

 _expression_. **BoxLinks**( ** _Style_**, ** _ShowArrows_**, ** _ShowLabels_**, ** _ColorMode_**, ** _CriticalColor_**, ** _NoncriticalColor_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Style_|Optional|**Long**|Specifies the style of link lines. Can be one of the following  **[PjLinkStyle](pjlinkstyle-enumeration-project.md)** constants: **pjLinkStraight** or **pjLinkRectilinear**.|
| _ShowArrows_|Optional|**Boolean**|**True** if link lines have arrows showing the direction of the link; otherwise, **False**.|
| _ShowLabels_|Optional|**Boolean**|**True** if link lines have labels showing the link type (FS, SS, SF, or FF); otherwise, **False**.|
| _ColorMode_|Optional|**Long**|Specifies how the color of link lines is determined. Can be one of the  **[PjLinkColorMode](pjlinkcolormode-enumeration-project.md)** constants.|
| _CriticalColor_|Optional|**Long**|The color of link lines between critical tasks. The default value is  **pjRed**. Can be one of the **[PjColor](pjcolor-enumeration-project.md)** constants.|
| _NoncriticalColor_|Optional|**Long**| The color of link lines between noncritical tasks. Can be one of the **[PjColor](pjcolor-enumeration-project.md)** constants. The default value is **pjBlack**.|

### Return Value

 **Boolean**


## Remarks

If no arguments are specified, the  **BoxLinks** method has no effect. If _ColorMode_ is **pjColorModePredecessor**, the _NoncriticalColor_ and _CriticalColor_ parameters are ignored.

To edit box link lines where the colors can be RGB values, use the  **[BoxLinksEx](application-boxlinksex-method-project.md)** method.


## Example

The following example shows link labels and then sets critical links to a purple color and noncritical links to a teal color.


```vb
Sub BoxLink_ChangeColor() 
 
 'Activate the Network Diagram view 
 ViewApply Name:="Network Diagram" 
 
 BoxLinks Style:=ShowLabels:=True, ColorMode:=pjColorModeCustom, _ 
 CriticalColor:=pjPurple, NoncriticalColor:=pjTeal 
End Sub
```


