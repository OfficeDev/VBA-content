---
title: ApplicationSettings.BIDITextUI Property (Visio)
keywords: vis_sdr.chm16260025
f1_keywords:
- vis_sdr.chm16260025
ms.prod: visio
api_name:
- Visio.ApplicationSettings.BIDITextUI
ms.assetid: a358e155-9ba0-42ca-3192-3fc90ee19559
ms.date: 06/08/2017
---


# ApplicationSettings.BIDITextUI Property (Visio)

Gets the current setting for display of right-to-left languages. Read-only.


## Syntax

 _expression_ . **BIDITextUI**

 _expression_ An expression that returns a **ApplicationSettings** object.


### Return Value

VisRegionalUIOptions


## Remarks

The following  **VisRegionalUIOptions** constants, which are declared in the Visio type libary, show the possible values for the **BIDITextUI** property.



|**Constant**|**Value**|**Description**|
|:-----|:-----|:-----|
| **visRegionalUIOptionsHide**|0|Always hides regional UI.|
| **visRegionalUIOptionsShow**|1|Always shows regional UI|
The setting of the  **BIDITextUI** property corresponds to the regional options setting in the **Microsoft Office Language Preferences** dialog box. (Click **Start**, point to  **All Programs**, point to  **Microsoft Office**, point to  **Microsoft Office Tools**, and then click  **Microsoft Office Language Preferences**. 

The setting of the  **BIDITextUI** property influences the setting of the **[ApplicationSettings.ComplexTextUI](applicationsettings-complextextui-property-visio.md)** property. If **BIDITextUI** is set to **visRegionalUIOptionsShow** , **ComplexTextUI** is set to that value as well.


