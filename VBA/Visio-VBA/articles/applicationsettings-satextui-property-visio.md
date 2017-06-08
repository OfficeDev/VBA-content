---
title: ApplicationSettings.SATextUI Property (Visio)
keywords: vis_sdr.chm16260020
f1_keywords:
- vis_sdr.chm16260020
ms.prod: visio
api_name:
- Visio.ApplicationSettings.SATextUI
ms.assetid: e8bdb2bd-a54b-01e4-8ee7-c3d5c3156854
ms.date: 06/08/2017
---


# ApplicationSettings.SATextUI Property (Visio)

Gets the current setting for display of South Asian languages. Read-only.


## Syntax

 _expression_ . **SATextUI**

 _expression_ An expression that returns a **ApplicationSettings** object.


### Return Value

VisRegionalUIOptions


## Remarks

The following  **VisRegionalUIOptions** constants, which are declared in the Visio type libary, show the possible values for the **SATextUI** property.



|**Constant**|**Value**|**Description**|
|:-----|:-----|:-----|
| **visRegionalUIOptionsHide**|0|Always hides regional UI.|
| **visRegionalUIOptionsShow**|1|Always shows regional UI|
The setting of the  **SATextUI** property corresponds to the regional options setting in the **Microsoft Office Language Preferences** dialog box. (Click **Start**, point to  **All Programs**, point to  **Microsoft Office**, point to  **Microsoft Office Tools**, and then click  **Microsoft Office Language Preferences**). 

The setting of the  **SATextUI** property influences the setting of the **[ApplicationSettings.ComplexTextUI](applicationsettings-complextextui-property-visio.md)** property. If **SATextUI** is set to **visRegionalUIOptionsShow** , **ComplexTextUI** is set to that value as well.


