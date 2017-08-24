---
title: TextRange.InsertPageNumber Method (Publisher)
keywords: vbapb10.chm5308486
f1_keywords:
- vbapb10.chm5308486
ms.prod: publisher
api_name:
- Publisher.TextRange.InsertPageNumber
ms.assetid: f71d3b40-0263-93fa-d7e3-d815b90f71f7
ms.date: 06/08/2017
---


# TextRange.InsertPageNumber Method (Publisher)

Returns a  **[TextRange](textrange-object-publisher.md)** object that represents a page number field in a publication.


## Syntax

 _expression_. **InsertPageNumber**( **_Type_**)

 _expression_A variable that represents a  **TextRange** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|Type|Optional| **PbPageNumberType**|Specifies whether the page number is the current page number or the next or previous page number of a linked text box.|

### Return Value

TextRange


## Remarks

Type can be one of these  **PbPageNumberType** constants.



|**Constant**|**Description**|
|:-----|:-----|
| **pbPageNumberCurrent**|The default.|
| **pbPageNumberNextInStory**|Inserts the page number of the next linked text box.|
| **pbPageNumberPreviousInStory**|Inserts the page number of the previous linked text box.|

## Example

This example inserts a page number field in a shape on the master page so that the current page number appears at the top of each page.


```vb
Sub PageNumberShape() 
 With ActiveDocument.MasterPages(1).Shapes _ 
 .AddShape(Type:=msoShape5pointStar, Left:=36, _ 
 Top:=36, Width:=50, Height:=50) 
 With .TextFrame.TextRange 
 .InsertPageNumber 
 .ParagraphFormat.Alignment = pbParagraphAlignmentCenter 
 End With 
 .Fill.ForeColor.RGB = RGB(Red:=125, Green:=125, Blue:=255) 
 End With 
End Sub
```


