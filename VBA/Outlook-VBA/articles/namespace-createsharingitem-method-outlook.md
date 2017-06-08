---
title: NameSpace.CreateSharingItem Method (Outlook)
keywords: vbaol11.chm790
f1_keywords:
- vbaol11.chm790
ms.prod: outlook
api_name:
- Outlook.NameSpace.CreateSharingItem
ms.assetid: 4c93d347-cc39-eb5d-bf08-125b69f91eb6
ms.date: 06/08/2017
---


# NameSpace.CreateSharingItem Method (Outlook)

Creates a new  **[SharingItem](sharingitem-object-outlook.md)** object.


## Syntax

 _expression_ . **CreateSharingItem**( **_Context_** , **_Provider_** )

 _expression_ An expression that returns a **NameSpace** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Context_|Required| **Variant**|Either a  **String** value or a **[Folder](folder-object-outlook.md)** object representing the sharing context to be used.|
| _Provider_|Optional| **Variant**|An  **[OlSharingProvider](olsharingprovider-enumeration-outlook.md)** value representing the sharing provider to be used.|

### Return Value

A  **SharingItem** object that represents a sharing message for the specified context.


## Remarks

If a  **String** value is specified in _Context_ , the method assumes that a URL has been provided as a sharing context. If a **[Folder](folder-object-outlook.md)** object is specified in _Context_ , the method attempts to discover the sharing context from the folder. If no sharing context exists, or if more than one sharing context exists, an error occurs.

If  _Provider_ is not specified, the method attempts to use the appropriate sharing provider for the value specified in _Context_ .


## See also


#### Concepts


[NameSpace Object](namespace-object-outlook.md)

