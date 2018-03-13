---
title: Document.RemoveHiddenInformation Method (Visio)
keywords: vis_sdr.chm10500000
f1_keywords:
- vis_sdr.chm10500000
ms.prod: visio
api_name:
- Visio.Document.RemoveHiddenInformation
ms.assetid: cc097f8b-5e74-9b44-4ba9-19537169c88b
ms.date: 06/08/2017
---


# Document.RemoveHiddenInformation Method (Visio)

Removes hidden information, such as personal information and external data, from a Microsoft Visio document.


## Syntax

 _expression_ . **RemoveHiddenInformation**( **_VisRemoveHiddenInfoItems_** )

 _expression_ An expression that returns a **Document** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _VisRemoveHiddenInfoItems_|Required| **Long**|The items to be removed. A combination of one or more enumerated values from the  **VisRemoveHiddenInfoItems** enumeration. See Remarks for possible values.|

### Return Value

Nothing


## Remarks

For the  _VisRemoveHiddenInfoItems_ parameter, pass a combination of one or more of the following values from the **VisRemoveHiddenInfoItems** enumeration, which is declared in the Visio type library.



| <strong>Constant</strong>              | <strong>Value</strong> | <strong>Description</strong>                              |
|:---------------------------------------|:-----------------------|:----------------------------------------------------------|
| <strong>visRHIPersonalInfo</strong>    | 1                      | Removes personal information.                             |
| <strong>visRHIDataRecordsets</strong>  | 16                     | Removes data recordsets.                                  |
| <strong>visRHIPreview</strong>         | 2                      | Removes document preview thumbnail images.                |
| <strong>visRHIMasters</strong>         | 4                      | Removes unused masters.                                   |
| <strong>visRHIStyles</strong>          | 8                      | Removes unused styles, themes, and other display formats. |
| <strong>visRHIValidationRules</strong> | 32                     | Removes data rows not linked to shapes in the drawing.    |

Calling the  **RemoveHiddenInformation** method is the equivalent of the selecting the options available in the **Remove Hidden Information** dialog box (click the **File** tab, click **Info**, and then click  **Remove Personal Information**).


## Example

This Microsoft Visual Basic for Applications (VBA) macro shows how to use the  **RemoveHiddenInformation** method to remove personal information and the preview thumbnail from the active document.


```vb
Public Sub RemoveHiddenInformation_Example() 

    ActiveDocument.RemoveHiddenInformation visRHIPersonalInfo + visRHIPreview 

End Sub
```


