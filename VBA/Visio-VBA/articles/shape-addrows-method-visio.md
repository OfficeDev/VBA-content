---
title: Shape.AddRows Method (Visio)
keywords: vis_sdr.chm11216055
f1_keywords:
- vis_sdr.chm11216055
ms.prod: visio
api_name:
- Visio.Shape.AddRows
ms.assetid: 8b267f98-e077-0854-a1aa-a0ce8719a2c5
ms.date: 06/08/2017
---


# Shape.AddRows Method (Visio)

Adds the specified number of rows to a ShapeSheet section at a specified position.


## Syntax

 _expression_ . **AddRows**( **_Section_** , **_Row_** , **_RowTag_** , **_RowCount_** )

 _expression_ A variable that represents a **Shape** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Section_|Required| **Integer**|The section in which to add the rows.|
| _Row_|Required| **Integer**|The position at which to add the rows.|
| _RowTag_|Required| **Integer**|The type of rows to add.|
| _RowCount_|Required| **Integer**|The number of rows to add.|

### Return Value

Integer


## Remarks

If the ShapeSheet section does not exist, the  **AddRows** method creates a section that has blank rows. New cells in new rows are initialized with default formulas, if applicable. Otherwise, a program must include statements to set the formulas for the new cells. An error occurs if the row cannot be added.

The Visio type library declares the constants for  _RowTag_ in **[VisRowIndices](visrowindices-enumeration-visio.md)** .

The row constants declared by the Visio type library serve as base positions at which a section's rows begin. Add offsets to these constants to specify the first row and beyond, for example,  **visRowFirst** +0, **visRowFirst** +1, and so on. To add rows at the end of a section, pass the constant **visRowLast** for the _Row_ argument. The value returned is the actual row index.

The  _RowTag_ argument specifies the type of rows to add. Pass **visTagDefault** (0) as the _RowTag_ argument to generate a section's default row type. Explicit tags are useful when adding rows to Geometry, Connection Points, and Controls sections. See the **RowType** property for descriptions of valid row types for these sections. Passing an invalid row type generates an error.

If you try to add rows to a Character, Tabs, or Paragraph section, an error occurs.

The  **AddRows** method cannot add named rows. To add named rows, use the **AddNamedRow** method.

If you add rows to a section that has nameable rows (for example, the Connection Points or Controls section), the  _Row_ argument is ignored. By default, named rows are named in the order added, for example, Row_1, Row_2, and so forth. Naming order is influenced, however, by any existing rows or previously deleted rows.


