---
title: TabStops2.Add Method (Office)
ms.prod: office
api_name:
- Office.TabStops2.Add
ms.assetid: 850b5a3d-c85e-33e5-b8d5-8ca469632e39
ms.date: 06/08/2017
---


# TabStops2.Add Method (Office)

Adds a new tab stop to the specified  **TabStops2** object.


## Syntax

 _expression_. **Add**( **_Type_**, **_Position_** )

 _expression_ An expression that returns a **TabStops2** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Type_|Required|**MsoTabStopType**|The type of tab stop to add.|
| _Position_|Required|**Single**|The horizontal position of the new tab stop relative to the left edge of the text frame. Numeric values are evaluated in points; strings are evaluated in the units specified and can be in any measurement unit supported by the Microsoft Office product. |

### Return Value

TabStop2


## Remarks

Examples of  **MsoTabStopType** types include **msoTabStopCenter**, **msoTabStopLeft**, and **msoTabStopRight**.


## See also


#### Concepts


[TabStops2 Object](tabstops2-object-office.md)
#### Other resources


[TabStops2 Object Members](tabstops2-members-office.md)

