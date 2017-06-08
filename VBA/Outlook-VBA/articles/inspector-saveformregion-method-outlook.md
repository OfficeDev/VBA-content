---
title: Inspector.SaveFormRegion Method (Outlook)
keywords: vbaol11.chm2983
f1_keywords:
- vbaol11.chm2983
ms.prod: outlook
api_name:
- Outlook.Inspector.SaveFormRegion
ms.assetid: 8ed73f85-3f6e-11bb-cc6f-c5c2668e5eb2
ms.date: 06/08/2017
---


# Inspector.SaveFormRegion Method (Outlook)

Saves the specified page in design mode in the inspector to the specified file.


## Syntax

 _expression_ . **SaveFormRegion**( **_Page_** , **_FileName_** )

 _expression_ A variable that represents an **Inspector** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Page_|Required| **Object**|The page displaying the form region in the inspector.|
| _FileName_|Required| **String**|The full local file path to an Outlook Form Storage (.OFS) file that the form region is being saved to. |

## Remarks

In order for  **SaveFormRegion** to save the form region, the form region must have been loaded in design mode in the inspector.


## See also


#### Concepts


[Inspector Object](inspector-object-outlook.md)

