---
title: Shape.AddRow Method (Visio)
keywords: vis_sdr.chm11216050
f1_keywords:
- vis_sdr.chm11216050
ms.prod: visio
api_name:
- Visio.Shape.AddRow
ms.assetid: 8b8dcf65-9b42-b3bf-0da3-61d3fbd02996
ms.date: 06/08/2017
---


# Shape.AddRow Method (Visio)

Adds a row to a ShapeSheet section at a specified position.


## Syntax

 _expression_ . **AddRow**( **_Section_** , **_Row_** , **_RowTag_** )

 _expression_ A variable that represents a **Shape** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Section_|Required| **Integer**|The section in which to add the row.|
| _Row_|Required| **Integer**| The position at which to add the row.|
| _RowTag_|Required| **Integer**|The type of row to add.|

### Return Value

Integer


## Remarks

If the ShapeSheet section does not already exist, it is created with a blank row. New cells in new rows are initialized with default formulas, if applicable. Otherwise, a program must include statements to set the formulas for the new cells. If the new row cannot be added, an error is generated.

The Visio type library declares row constants prefixed with  **visRow** in **[VisRowIndices](visrowindices-enumeration-visio.md)** .

Constants for rows in the Geometry, Connection Points, and Controls sections are prefixed with  **visTag** and declared by the type library in **[VisRowTags](visrowtags-enumeration-visio.md)** . To see a list of these constants, see the **RowType** property.

The row constants declared by the Visio type library serve as base positions at which a section's rows begin. Add offsets to these constants to specify the first row and beyond, for example,  **visRowFirst** +0, **visRowFirst** +1, and so on. To add rows at the end of a section, pass the constant **visRowLast** for the _Row_ argument. The value returned is the actual row index.

The  _RowTag_ argument specifies the type of row to add. To generate a section's default row type, pass **visTagDefault** (0) as the _RowTag_ argument. Explicit tags are useful when adding rows to Geometry, Connection Points, and Controls sections. See the **RowType** property for descriptions of valid row types for these sections. Passing an invalid row type generates an error.

If you try to add a row to a Character, Tabs, or Paragraph section, an error occurs.

The  **AddRow** method cannot add named rows. To add named rows, use the **AddNamedRow** method.

If you add rows to a section that has nameable rows (for example, the Connection Points or Controls section), the  _Row_ argument is ignored. By default, named rows are named in the order added, for example, Row_1, Row_2, and so forth. Naming order is influenced, however, by any existing rows or previously deleted rows.


## Example

The following example shows how to add a section to a ShapeSheet and how to add a row to the section at a specified position.


```vb
 
Public Sub AddRow_Example() 
 
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
 
 'Add a scratch section to the ShapeSheet. 
 vsoShape.AddSection visSectionScratch 
 
 'Add a row to the scratch section. 
 vsoShape.AddRow visSectionScratch, visRowScratch, 0 
 
End Sub
```


