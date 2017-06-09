---
title: ContainerProperties.GetMemberShapes Method (Visio)
keywords: vis_sdr.chm17662350
f1_keywords:
- vis_sdr.chm17662350
ms.prod: visio
api_name:
- Visio.ContainerProperties.GetMemberShapes
ms.assetid: 4fb246c7-b86d-4e90-ef91-9cac988dbbb8
ms.date: 06/08/2017
---


# ContainerProperties.GetMemberShapes Method (Visio)

Returns the shape identifiers (IDs) of all members of the container, as specified.


## Syntax

 _expression_ . **GetMemberShapes**( **_ContainerFlags_** )

 _expression_ A variable that represents a **[ContainerProperties](containerproperties-object-visio.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _ContainerFlags_|Required| **Long**|Specifies which container member shape IDs to return.|

### Return Value

 **Long()**


## Remarks

The _ContainerFlags_ parameter can be one or more of the following **VisContainerFlags** constants.



|**Constant**|**Value**|**Description**|
|:-----|:-----|:-----|
| **visContainerFlagsDefault**|0|Returns all shape types and includes items in nested containers.|
| **visContainerFlagsExcludeContainers**|1|Excludes member shapes that are containers.|
| **visContainerFlagsExcludeConnectors**|2|Excludes member shapes that are connectors.|
| **visContainerFlagsExcludeCallouts**|4|Excludes member shapes that are callouts.|
| **visContainerFlagsExcludeElements**|8|Excludes member shapes that are not containers, lists, connectors, or callouts.|
| **visContainerFlagsExcludeNested**|16|Excludes any member shapes that are members of containers or lists nested within the container.|
| **visContainerFlagsExcludeListMembers**|32|Excludes members of a list container that are explicitly members of any list. Does not exclude other shapes in the list container.|
 **GetMemberShapes** returns an empty array if there are no member shapes.


## Example

The following Visual Basic for Applications (VBA) example shows how to use the  **GetMemberShapes** method to get the IDs of all member shapes in a specified container on the active page, loop through those shapes, and print the ID of each shape in the **Immediate** window.


```vb
For Each memberID In vsoContainerShape.ContainerProperties.GetMemberShapes(visContainerFlagsDefault) 
    Set vsoShape = ActivePage.Shapes.ItemFromID(memberID) 
    Debug.Print vsoShape.ID
Next
```


