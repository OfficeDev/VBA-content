---
title: ApplicationSettings.AsianTextUI Property (Visio)
keywords: vis_sdr.chm16251920
f1_keywords:
- vis_sdr.chm16251920
ms.prod: visio
api_name:
- Visio.ApplicationSettings.AsianTextUI
ms.assetid: b317afda-5014-6c53-44e1-a713dabee111
ms.date: 06/08/2017
---


# ApplicationSettings.AsianTextUI Property (Visio)

Gets whether Asian text is displayed in the Microsoft Visio user interface. Read-only.


## Syntax

 _expression_ . **AsianTextUI**

 _expression_ An expression that returns a **ApplicationSettings** object.


### Return Value

VisRegionalUIOptions


## Remarks

The following  **VisRegionalUIOptions** constants, which are declared in the Visio type libary, show the possible values for the **AsianTextUI** property.



|**Constant**|**Value**|**Description**|
|:-----|:-----|:-----|
| **visRegionalUIOptionsHide**|0|Always hides regional UI.|
| **visRegionalUIOptionsShow**|1|Always shows regional UI.|
You can change language settings in the  **Microsoft Office Language Preferences** dialog box. (Click **Start**, point to  **All Programs**, point to  **Microsoft Office**, point to  **Microsoft Office Tools**, and then click  **Microsoft Office Language Preferences**.


