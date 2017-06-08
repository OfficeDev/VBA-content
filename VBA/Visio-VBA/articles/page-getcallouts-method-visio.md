---
title: Page.GetCallouts Method (Visio)
keywords: vis_sdr.chm10962170
f1_keywords:
- vis_sdr.chm10962170
ms.prod: visio
api_name:
- Visio.Page.GetCallouts
ms.assetid: a0300c64-4bdd-e442-c00c-a727debbf6b8
ms.date: 06/08/2017
---


# Page.GetCallouts Method (Visio)

Returns the list of identifiers of the callout shapes on the page.


## Syntax

 _expression_ . **GetCallouts**( **_NestedOptions_** )

 _expression_ A variable that represents a **[Page](page-object-visio.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _NestedOptions_|Required| **[VisContainerNested](viscontainernested-enumeration-visio.md)**|A constant that indicates whether to exclude shapes on the page that are contained by containers or lists. See Remarks for possible values.|

### Return Value

 **Long()**


## Remarks

The  _NestedOptions_ parameter must be one of the following **VisContainerNested** constants.



|**Constant**|**Value**|**Description**|
|:-----|:-----|:-----|
| **visContainerIncludeNested**|0|Include shapes that are in nested containers.|
| **visContainerExcludeNested**|1|Exclude shapes that are in nested containers.|

