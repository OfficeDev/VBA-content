
# TabStops.Add Method (PowerPoint)

 **Last modified:** July 28, 2015

Creates a tab stop and adds it to the  **TabStops** collection.

## Syntax

 _expression_. **Add**( **_Type_**,  **_Position_**)

 _expression_A variable that represents a  **TabStops** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|Type|Required| **PpTabStopType**|The type of the tab stop to be added.|
|Position|Required| **Single**|The position of the tab stop in the tab stops collection.|

### Return Value

TabStop


## Remarks

The  _Type_ parameter value can be one of these **PpTabStopType** constants.



|**Constant**|**Description**|
|:-----|:-----|
| **ppTabStopCenter**|Center tab stop.|
| **ppTabStopDecimal**|Decimal tab stop.|
| **ppTabStopLeft**|Left tab stop.|
| **ppTabStopMixed**|Mixed tab stop.|
| **ppTabStopRight**|Right tab stop.|

## See also


#### Concepts


 [TabStops Object](e23b36de-6a4d-84e5-bec1-8c3e0fd80c13.md)
#### Other resources


 [TabStops Object Members](62f6b7f4-45f8-108c-6294-8f24d3b2058c.md)
