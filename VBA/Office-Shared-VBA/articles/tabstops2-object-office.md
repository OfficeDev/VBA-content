---
title: TabStops2 Object (Office)
ms.prod: office
api_name:
- Office.TabStops2
ms.assetid: 1d1d8054-19eb-cd65-f37d-36e93e7fc347
ms.date: 06/08/2017
---


# TabStops2 Object (Office)

The collection of  **TabStop2** objects.


## Remarks

Tab stops are indexed numerically from left to right along the ruler.


## Example

 The following example removes the first custom tab stop from the first paragraph in the active Microsoft Publisher publication.


```
Sub ClearTabStop() 
    ActiveDocument.Pages(1).Shapes(1).TextFrame.TextRange _ 
        .ParagraphFormat.Tabs(1).Clear 
End Sub 

```


## Methods



|**Name**|
|:-----|
|[Add](tabstops2-add-method-office.md)|
|[Item](tabstops2-item-method-office.md)|

## Properties



|**Name**|
|:-----|
|[Application](tabstops2-application-property-office.md)|
|[Count](tabstops2-count-property-office.md)|
|[Creator](tabstops2-creator-property-office.md)|
|[DefaultSpacing](tabstops2-defaultspacing-property-office.md)|
|[Parent](tabstops2-parent-property-office.md)|

## See also


#### Other resources


[Object Model Reference](http://msdn.microsoft.com/library/499c789a-aba2-0fad-649a-0ea964cd3b5e%28Office.15%29.aspx)
