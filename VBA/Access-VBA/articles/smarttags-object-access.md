---
title: SmartTags Object (Access)
keywords: vbaac10.chm13280
f1_keywords:
- vbaac10.chm13280
ms.prod: access
api_name:
- Access.SmartTags
ms.assetid: 79c0e84e-e0a1-35b8-b826-9d2cde3bd485
ms.date: 06/08/2017
---


# SmartTags Object (Access)

Represents the collection of smart tags for a control on a form, report, or data access page.


## Remarks

To return a single  **[SmartTag](smarttag-object-access.md)** object, use the **Item** method or use **SmartTags** ( _Index_), where  _Index_ represents the number of the smart tag.


 **Note**  Unlike the  **SmartTags** collections in Microsoft Excel and Microsoft Word, the **SmartTags** collection in Microsoft Access is zero-based. Therefore, the code `control.SmartTags(0)` returns the first smart tag for the specified control.


## Methods



|**Name**|
|:-----|
|[Add](smarttags-add-method-access.md)|

## Properties



|**Name**|
|:-----|
|[Application](smarttags-application-property-access.md)|
|[Count](smarttags-count-property-access.md)|
|[Item](smarttags-item-property-access.md)|
|[Parent](smarttags-parent-property-access.md)|

## See also


#### Other resources


[Access Object Model Reference](http://msdn.microsoft.com/library/2de134a4-6c5c-d2a3-8377-f4dd973ba650%28Office.15%29.aspx)
