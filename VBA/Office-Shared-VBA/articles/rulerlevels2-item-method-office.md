---
title: RulerLevels2.Item Method (Office)
ms.prod: office
api_name:
- Office.RulerLevels2.Item
ms.assetid: b6791181-ea32-62e3-3b9a-1b60f436bc91
ms.date: 06/08/2017
---


# RulerLevels2.Item Method (Office)

Gets a member of the  **RulerLevels2** collection.


## Syntax

 _expression_. **Item**( **_Index_** )

 _expression_ An expression that returns a **RulerLevels2** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Index_|Required|**Variant**|The index number of the object to be returned.|

### Return Value

RulerLevel2


## Example

This example sets the first-line indent and the hanging indent for outline level one in body text on the slide master for the active presentation.


```
With ActivePresentation.SlideMaster.TextStyles.Item(ppBodyStyle) 
 With .Ruler2.Levels.Item(1) ' sets indents for level 1 
 .FirstMargin = 9 
 .LeftMargin = 54 
 End With 
End With 

```


## See also


#### Concepts


[RulerLevels2 Object](rulerlevels2-object-office.md)
#### Other resources


[RulerLevels2 Object Members](rulerlevels2-members-office.md)

