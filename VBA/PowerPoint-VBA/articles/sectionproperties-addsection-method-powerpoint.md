---
title: SectionProperties.AddSection Method (PowerPoint)
keywords: vbapp10.chm725009
f1_keywords:
- vbapp10.chm725009
ms.prod: powerpoint
api_name:
- PowerPoint.SectionProperties.AddSection
ms.assetid: bdad42c8-0d2b-91cc-67c5-452abe28d658
ms.date: 06/08/2017
---


# SectionProperties.AddSection Method (PowerPoint)

Adds a new section at the specified index position and returns the index of the newly created section.


## Syntax

 _expression_. **AddSection**( **_sectionIndex_**, **_sectionName_** )

 _expression_ A variable that represents a **SectionProperties** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _sectionIndex_|Required|**Integer**|The index of the section before which to add the section.|
| _sectionName_|Optional|**Variant**|The name of the new section.|

### Return Value

Integer


## Remarks

If a section already exists at the specified index position, calling  **AddSection** is equivalent to adding an empty section immediately before that section.

The indices of sections after the newly inserted section are automatically incremented by one.

sectionIndex can be one larger than the total number of existing sections, as long as it is less than the maximum number of sections allowed (512); this creates a new, empty section at the end.

If no sections exist in the presentation, calling this method and passing a sectionIndex value of 1 creates the first section.


## See also


#### Concepts


[SectionProperties Object](sectionproperties-object-powerpoint.md)

