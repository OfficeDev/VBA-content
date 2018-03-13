---
title: Selection.SetContainerFormat Method (Visio)
keywords: vis_sdr.chm11162235
f1_keywords:
- vis_sdr.chm11162235
ms.prod: visio
api_name:
- Visio.Selection.SetContainerFormat
ms.assetid: b0766138-07da-4539-b254-7692529e0771
ms.date: 06/08/2017
---


# Selection.SetContainerFormat Method (Visio)

Changes the formatting of one aspect of all the containers in the selection, and returns an array of identifiers of shapes that belong to the containers and whose formatting was changed. 


## Syntax

 _expression_ . **SetContainerFormat**( **_FormatType_** , **_[FormatValue]_** )

 _expression_ An expression that returns a **[Selection](selection-object-visio.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _FormatType_|Required| **[VisContainerFormatType](viscontainerformattype-enumeration-visio.md)**|The container formatting to change. See Remarks for possible values.|
| _FormatValue_|Optional| **Variant**|The new format to apply.|

### Return Value

 **Long()**


## Remarks

The  _FormatType_ parameter must be one of the following **VisContainerFormatType** constants.



| <strong>Constant</strong>                              | <strong>Value</strong> | <strong>Description</strong>                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                           |
|:-------------------------------------------------------|:-----------------------|:---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| <strong>visContainerFormatLockMembership</strong>      | 0                      | Apply one of the  <strong><a href="containerproperties-lockmembership-property-visio.md" data-raw-source="[LockMembership](containerproperties-lockmembership-property-visio.md)">LockMembership</a></strong> property values, as specified in <em>FormatValue</em>.  <em>FormatValue</em> is required, and must be of type <strong>Boolean</strong> (preferred) or or another type that can be converted to <strong>Boolean</strong> .                                                                                                                                                                                                                                                                                                |
| <strong>visContainerFormatContainerAutoResize</strong> | 1                      | Apply one of the  <strong><a href="containerproperties-resizeasneeded-property-visio.md" data-raw-source="[ResizeAsNeeded](containerproperties-resizeasneeded-property-visio.md)">ResizeAsNeeded</a></strong> property values, as specified in <em>FormatValue</em>. Applies to normal containers only.  <em>FormatValue</em> is required, must be of type <strong>Short</strong> (preferred) or of another type that can be converted to <strong>Short</strong> , and must be equal to a constant in the range of those in the <strong><a href="viscontainerautoresize-enumeration-visio.md" data-raw-source="[VisContainerAutoResize](viscontainerautoresize-enumeration-visio.md)">VisContainerAutoResize</a></strong> enumeration. |
| <strong>visContainerFormatFitToContents</strong>       | 2                      | Fit contents to the container.  <em>FormatValue</em> is ignored.                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                       |

If the selection does not include any containers, this method has no effect.

If  _FormatType_ is of an incorrect type or is out of the range of **VisContainerFormatType** , Microsoft Visio returns an Invalid Parameter error.


