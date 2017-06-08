---
title: Shape.AddNamedRow Method (Visio)
keywords: vis_sdr.chm11216045
f1_keywords:
- vis_sdr.chm11216045
ms.prod: visio
api_name:
- Visio.Shape.AddNamedRow
ms.assetid: c18380b1-418d-454f-3c90-fa4624291628
ms.date: 06/08/2017
---


# Shape.AddNamedRow Method (Visio)

Adds a row that has the specified name to the specified ShapeSheet section.


## Syntax

 _expression_ . **AddNamedRow**( **_Section_** , **_RowName_** , **_RowTag_** )

 _expression_ A variable that represents a **Shape** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Section_|Required| **Integer**| The section in which the row is to be added.|
| _RowName_|Required| **String**|The name of the new row.|
| _RowTag_|Required| **Integer**| The type of row to be added.|

### Return Value

Integer


## Remarks

You can add named rows to the Actions ( **visSectionAction** ), Controls ( **visSectionControls** ), Shape Data ( **visSectionProp** ), User-Defined Cells ( **visSectionUser** ), Hyperlinks ( **visSectionHyperlink** ), and Connection Points ( **visSectionConnectionPts** ) ShapeSheet sections. You can access cells in the new rows by passing the row number returned by the **AddNamedRow** method to the **CellsSRC** property. Alternatively, you can access cells in the new rows by using the row's name with the **Cells** property. For details about cell references and cells in named rows, see the Actions, Controls, User-defined Cells, Hyperlink, Shape Data, or Connection Points row topics in the Microsoft Visio ShapeSheet Reference.

An empty row name string ("") creates a row with a default name.

Passing a value of  **visTagDefault** (0) in the _RowTag_ argument generates the default row type for the section. Explicit tags are useful when adding rows to a Connection Points section. See the **RowType** property for descriptions of valid row types for each section. Passing an invalid row type generates an error.

Adding a named row to a Connection Points section automatically converts any existing unnamed rows in the section into named rows, assigning them default names (Row_1, Row_2, and so on).


