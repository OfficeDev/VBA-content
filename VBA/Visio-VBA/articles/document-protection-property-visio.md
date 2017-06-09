---
title: Document.Protection Property (Visio)
keywords: vis_sdr.chm10550785
f1_keywords:
- vis_sdr.chm10550785
ms.prod: visio
api_name:
- Visio.Document.Protection
ms.assetid: f80cd284-e0e3-0663-c505-88311ffc9d3b
ms.date: 06/08/2017
---


# Document.Protection Property (Visio)

Determines how a document is protected from user customization. Read/write.


## Syntax

 _expression_ . **Protection**( **_bstrPassword_** )

 _expression_ A variable that represents a **Document** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _bstrPassword_|Optional| **Variant**|Not used.|

### Return Value

VisProtection


## Remarks

Beginning with Microsoft Office Visio 2003, the  **Protection** property ignores the _bstrPassword_ argument both when you get and when you set the value of the property.

This property is the equivalent of checking the  **Styles**,  **Shapes**,  **Preview**,  **Backgrounds**, and  **Master shapes** boxes in the **Protect Document** dialog box (in the **Drawing Explorer**, right-click the drawing name, and then click  **Protect Document**). 

The value of the  **Protection** property can be a combination of the following **VisProtection** constants.



|**Constant **|**Value **|
|:-----|:-----|
| **visProtectNone**|&;H0|
| **visProtectStyles**|&;H1|
| **visProtectShapes**|&;H2|
| **visProtectMasters**|&;H4|
| **visProtectBackgrounds**|&;H8|
| **visProtectPreviews**|&;H10|

