---
title: Shape.DeleteRow Method (Visio)
keywords: vis_sdr.chm11216170
f1_keywords:
- vis_sdr.chm11216170
ms.prod: visio
api_name:
- Visio.Shape.DeleteRow
ms.assetid: 892ca523-679d-c707-4aba-e43c011cb718
ms.date: 06/08/2017
---


# Shape.DeleteRow Method (Visio)

Deletes a row from a section in a ShapeSheet spreadsheet.


## Syntax

 _expression_ . **DeleteRow**( **_Section_** , **_Row_** )

 _expression_ A variable that represents a **Shape** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Section_|Required| **Integer**|The index of the section that contains the row.|
| _Row_|Required| **Integer**|The index of the row to delete.|

### Return Value

Nothing


## Remarks

To remove one row at a time from a ShapeSheet section, use the  **DeleteRow** method. If the section has indexed rows, the rows following the deleted row shift position. If the row does not exist, nothing is deleted.

You should not delete rows that define fundamental characteristics of a shape, such as the 1-D Endpoints row ( **visRowXForm1D** ) or the component row ( **visRowComponent** ) or the MoveTo row ( **visRowVertex** + 0) in a Geometry section. You cannot delete rows from sections represented by **visSectionCharacter** , **visSectionParagraph** , and **visSectionTab** .


## Example

The following Microsoft Visual Basic for Applications (VBA) macro shows how to use the  **DeleteRow** method to delete a ShapeSheet row.


```vb
Public Sub DeleteRow_Example() 
 
 Dim vsoPage As Visio.Page 
 Dim vsoShape As Visio.Shape 
 
 'Get the active page. 
 Set vsoPage = ActivePage 
 
 'If there isn't an active page, set the Page object 
 'to the first page of the active document. 
 If vsoPage Is Nothing Then 
 Set vsoPage = ActiveDocument.Pages(1) 
 End If 
 
 'Draw a rectangle on the active page. 
 Set vsoShape = vsoPage.DrawRectangle(1, 5, 5, 1) 
 
 'Add a scratch section to the ShapeSheet of the rectangle. 
 vsoShape.AddSection visSectionScratch 
 
 'Add a row to the scratch section. 
 vsoShape.AddRow visSectionScratch, visRowScratch, 0 
 
 'Delete the row from the scratch section. 
 vsoShape.DeleteRow visSectionScratch, visRowScratch 
 
End Sub
```


