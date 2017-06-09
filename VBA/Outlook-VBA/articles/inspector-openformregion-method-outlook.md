---
title: Inspector.OpenFormRegion Method (Outlook)
keywords: vbaol11.chm2982
f1_keywords:
- vbaol11.chm2982
ms.prod: outlook
api_name:
- Outlook.Inspector.OpenFormRegion
ms.assetid: c574d034-6c8e-388b-f93f-cf899db24ae6
ms.date: 06/08/2017
---


# Inspector.OpenFormRegion Method (Outlook)

Opens a page in design mode in the inspector for the specified form region.


## Syntax

 _expression_ . **OpenFormRegion**( **_Path_** )

 _expression_ A variable that represents an **Inspector** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Path_|Required| **String**|A full local file path to the Outlook Form Storage (.OFS) file for the form region that is to be opened in the inspector.|

### Return Value

An  **Object** that represents the page displaying the form region in the inspector.


## Remarks

If the inspector is not already in design mode,  **OpenFormRegion** will put it in design mode.


## See also


#### Concepts


[Inspector Object](inspector-object-outlook.md)

