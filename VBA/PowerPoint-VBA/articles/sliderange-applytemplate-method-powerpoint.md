---
title: SlideRange.ApplyTemplate Method (PowerPoint)
keywords: vbapp10.chm532036
f1_keywords:
- vbapp10.chm532036
ms.prod: powerpoint
api_name:
- PowerPoint.SlideRange.ApplyTemplate
ms.assetid: 3bf6d3e0-bc37-00f3-868e-869f51c62ad3
ms.date: 06/08/2017
---


# SlideRange.ApplyTemplate Method (PowerPoint)

Applies a design template to the specified slide range.


## Syntax

 _expression_. **ApplyTemplate**( **_FileName_** )

 _expression_ A variable that represents a **SlideRange** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _FileName_|Required|**String**|Specifies the name of the design template.|

 **Note**  If you refer to an uninstalled presentation design template in a string, a run-time error is generated. The template is not installed automatically regardless of your  **[FeatureInstall](application-featureinstall-property-powerpoint.md)** property setting. To use the **ApplyTemplate** method for a template that is not currently installed, you first must install the additional design templates. To do so, install the Additional Design Templates for PowerPoint by running the Microsoft Office installation program (click **Add/Remove Programs** or **Programs and Features** in Windows Control Panel).


## See also


#### Concepts


[SlideRange Object](sliderange-object-powerpoint.md)

