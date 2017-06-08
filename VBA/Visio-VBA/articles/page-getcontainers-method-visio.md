---
title: Page.GetContainers Method (Visio)
keywords: vis_sdr.chm10962165
f1_keywords:
- vis_sdr.chm10962165
ms.prod: visio
api_name:
- Visio.Page.GetContainers
ms.assetid: 17d9365b-f9ac-85ba-e1cb-cd02ea1a2f22
ms.date: 06/08/2017
---


# Page.GetContainers Method (Visio)

Returns an array of shape identifiers (IDs) of the container shapes on the page.


## Syntax

 _expression_ . **GetContainers**( **_NestedOptions_** )

 _expression_ A variable that represents a **[Page](page-object-visio.md)** object.


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
| **visContainerExcludeNested**|1|Exclude shapes that are in nested containers.|
You can use the  **[Shapes.ItemFromID](shapes-itemfromid-property-visio.md)** property to get the actual shapes from the IDs returned by **GetContainers** .


## Example

The following Visual Basic for Applications (VBA) example shows how to use the  **GetContainers** method to get the IDs of all the containers on a page, loop through those containers, and print each container name in the **Immediate** window. The example includes nested containers.


```vb
For Each containerID In vsoPage.GetContainers(visContainerIncludeNested)
    Set vsoContainerShape = vsoPage.Shapes.ItemFromID(containerID)
    Debug.Print vsoContainerShape.NameU
Next
```


