---
title: TabStop2 Object (Office)
ms.prod: office
api_name:
- Office.TabStop2
ms.assetid: fee461a9-684b-e6c2-a74a-d0aa161d0d9c
ms.date: 06/08/2017
---


# TabStop2 Object (Office)

Represents a single tab stop. The  **TabStop2** object is a member of the **TabStops2** collection.


## Remarks

Tab stops are indexed numerically from left to right along the ruler.


## Example

The following example removes the first custom tab stop from the selected paragraphs.


```
Sub ClearTabStop() 
 Selection.TextRange.ParagraphFormat.Tabs(1).Clear 
End Sub 

```


## Methods



|**Name**|
|:-----|
|[Clear](tabstop2-clear-method-office.md)|

## Properties



|**Name**|
|:-----|
|[Application](tabstop2-application-property-office.md)|
|[Creator](tabstop2-creator-property-office.md)|
|[Parent](tabstop2-parent-property-office.md)|
|[Position](tabstop2-position-property-office.md)|
|[Type](tabstop2-type-property-office.md)|

## See also


#### Other resources


[Object Model Reference](http://msdn.microsoft.com/library/499c789a-aba2-0fad-649a-0ea964cd3b5e%28Office.15%29.aspx)
