---
title: MasterPages.FindByPageID Method (Publisher)
keywords: vbapb10.chm589830
f1_keywords:
- vbapb10.chm589830
ms.prod: publisher
api_name:
- Publisher.MasterPages.FindByPageID
ms.assetid: 2d05a2ae-853d-bc4c-bff8-0f3489627052
ms.date: 06/08/2017
---


# MasterPages.FindByPageID Method (Publisher)

Returns a  **[Page](page-object-publisher.md)** object that represents the page with the specified page ID number. Each page is automatically assigned a unique ID number when it is created. Use the **[PageID](page-pageid-property-publisher.md)** property to return a page's ID number.


## Syntax

 _expression_. **FindByPageID**( **_PageID_**)

 _expression_A variable that represents a  **MasterPages** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|PageID|Required| **Long**|Specifies the ID number of the page you want to return. Publisher assigns this number when the page is created.|

### Return Value

Page


## Remarks

Unlike the  **[PageIndex](page-pageindex-property-publisher.md)** property, the  **PageID** property of a **Page** object won't change when you add pages to or rearrange pages in the publication. Therefore, using the **FindByPageID** method with the page ID number can be a more reliable way to return a specific **Page** object from a **[Pages](pages-object-publisher.md)** collection than using the **Item**method with the page's index number.


## Example

This example demonstrates how to retrieve the unique ID number for a  **Page** object and then use this number to return that **Page** object from the **Pages** collection and add a new shape to the page.


```vb
Sub FindPage() 
 Dim lngPageID As Long 
 
 'Get page ID 
 lngPageID = ActiveDocument.Pages.Add(Count:=1, After:=1).PageID 
 
 'Use page ID to add a new shape to the page 
 ActiveDocument.Pages.FindByPageID(PageID:=lngPageID) _ 
 .Shapes.AddShape Type:=msoShape5pointStar, _ 
 Left:=200, Top:=72, Width:=50, Height:=50 
 
End Sub
```


