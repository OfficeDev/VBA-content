---
title: Section Object (Publisher)
keywords: vbapb10.chm7471103
f1_keywords:
- vbapb10.chm7471103
ms.prod: publisher
api_name:
- Publisher.Section
ms.assetid: 7e92a8de-ed66-564b-2657-cef0fc2392b8
ms.date: 06/08/2017
---


# Section Object (Publisher)

Represents a Section of a publication or document.
 


## Example

Use  **Sections**.Item(index) where index is the index number, to return a single **Section** object. The following example sets a **Section** object to the first section in the **Sections** collection of the active document.
 

 

```
Dim objSection As Section 
Set objSection = ActiveDocument.Sections.Item(1)
```

Use  **Sections**.Add(StartPageIndex) where StartPageIndex is the index number of the page, to return a new section added to a document. A "Permission denied." error will be returned if the page already contains a section head. The following example adds a Section object to the second page of the active document.
 

 



```
Dim objSection As Section 
Set objSection = ActiveDocument.Sections.Add(StartPageIndex:=2)
```


## Methods



|**Name**|
|:-----|
|[Delete](section-delete-method-publisher.md)|

## Properties



|**Name**|
|:-----|
|[Application](section-application-property-publisher.md)|
|[ContinueNumbersFromPreviousSection](section-continuenumbersfromprevioussection-property-publisher.md)|
|[PageNumberFormat](section-pagenumberformat-property-publisher.md)|
|[PageNumberStart](section-pagenumberstart-property-publisher.md)|
|[Parent](section-parent-property-publisher.md)|
|[ShowHeaderFooterOnFirstPage](section-showheaderfooteronfirstpage-property-publisher.md)|
|[StartPageIndex](section-startpageindex-property-publisher.md)|

