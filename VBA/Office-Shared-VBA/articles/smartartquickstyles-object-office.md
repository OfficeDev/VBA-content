---
title: SmartArtQuickStyles Object (Office)
ms.prod: office
api_name:
- Office.SmartArtQuickStyles
ms.assetid: d488ac12-160b-c518-2b56-cc0a3a45c6b7
ms.date: 06/08/2017
---


# SmartArtQuickStyles Object (Office)

Represents a collection of Smart Art quick styles.


## Example

The following code changes the quick style of a Smart Art diagram in Microsoft PowerPoint.


```
ActivePresentation.Slides(1).Shapes(1).SmartArt.QuickStyle = Application.SmartArtQuickStyles(i)
```


## Methods



|**Name**|
|:-----|
|[Item](smartartquickstyles-item-method-office.md)|

## Properties



|**Name**|
|:-----|
|[Application](smartartquickstyles-application-property-office.md)|
|[Count](smartartquickstyles-count-property-office.md)|
|[Creator](smartartquickstyles-creator-property-office.md)|
|[Parent](smartartquickstyles-parent-property-office.md)|

## See also


#### Other resources


[Object Model Reference](http://msdn.microsoft.com/library/499c789a-aba2-0fad-649a-0ea964cd3b5e%28Office.15%29.aspx)
