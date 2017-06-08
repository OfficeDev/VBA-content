---
title: Selection.GetContainers Method (Visio)
keywords: vis_sdr.chm11162165
f1_keywords:
- vis_sdr.chm11162165
ms.prod: visio
api_name:
- Visio.Selection.GetContainers
ms.assetid: 8e04bed5-f9ef-04bf-3013-c6dd623f9f63
ms.date: 06/08/2017
---


# Selection.GetContainers Method (Visio)

Returns an array of shape identifiers (IDs) of the container shapes in the selection.


## Syntax

 _expression_ . **GetContainers**( **_NestedOptions_** )

 _expression_ A variable that represents a **[Selection](selection-object-visio.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _NestedOptions_|Required| **[VisContainerNested](viscontainernested-enumeration-visio.md)**|Indicates whether to exclude shapes in nested containers. See Remarks for possible values.|

### Return Value

 **Long()**


## Remarks

The  _NestedOptions_ parameter must be one of the following **VisContainerNested** constants.



|**Constant**|**Value**|**Description**|
|:-----|:-----|:-----|
| **visContainerIncludeNested**|0|Include shapes that are in nested containers.|
| **visContainerExcludeNested**|1|Exclude shapes that are in nested containers..|
You can use the  **[Shapes.ItemFromID](shapes-itemfromid-property-visio.md)** property to get the actual shapes from the IDs returned by **GetContainers** .


