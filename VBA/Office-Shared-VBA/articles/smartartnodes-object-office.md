---
title: SmartArtNodes Object (Office)
ms.prod: office
api_name:
- Office.SmartArtNodes
ms.assetid: 4c35e5a4-15a1-dd6d-85a2-eb30cbaa3093
ms.date: 06/08/2017
---


# SmartArtNodes Object (Office)

Represents a collection of nodes within a Smart Art diagram. 


## Remarks

These nodes correspond directly to semantic elements contained within the data model of the graphic.


## Example

The following code returns the number of nodes in the Smart Art diagram.


```
ActivePresentation.Slides(1).Shapes(1).SmartArtNodes.Count
```


## Methods



|**Name**|
|:-----|
|[Add](smartartnodes-add-method-office.md)|
|[Item](smartartnodes-item-method-office.md)|

## Properties



|**Name**|
|:-----|
|[Application](smartartnodes-application-property-office.md)|
|[Count](smartartnodes-count-property-office.md)|
|[Creator](smartartnodes-creator-property-office.md)|
|[Parent](smartartnodes-parent-property-office.md)|

## See also


#### Other resources


[Object Model Reference](http://msdn.microsoft.com/library/499c789a-aba2-0fad-649a-0ea964cd3b5e%28Office.15%29.aspx)
