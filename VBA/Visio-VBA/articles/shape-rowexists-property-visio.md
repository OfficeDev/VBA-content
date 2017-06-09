---
title: Shape.RowExists Property (Visio)
keywords: vis_sdr.chm11214250
f1_keywords:
- vis_sdr.chm11214250
ms.prod: visio
api_name:
- Visio.Shape.RowExists
ms.assetid: bd89deb9-eda3-18d8-6305-bd380d9e649f
ms.date: 06/08/2017
---


# Shape.RowExists Property (Visio)

Determines whether a ShapeSheet row exists. Read-only.


## Syntax

 _expression_ . **RowExists**( **_Section_** , **_Row_** , **_fExistsLocally_** )

 _expression_ A variable that represents a **Shape** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Section_|Required| **Integer**|The row's section index.|
| _Row_|Required| **Integer**|The row's row index.|
| _fExistsLocally_|Required| **Integer**|The scope of the search.|

### Return Value

Integer


## Remarks

If  _fExistsLocally_ is **False** (0), the **RowExists** property returns **True** if the object either contains or inherits the specified row.

If  _fExistsLocally_ is **True** (non-zero), the **RowExists** property returns **True** only if the object contains the specified row locally; if the row is inherited, the **RowExists** property returns **False** .

For a list of row index values, see the  **AddRow** method or view the Visio type library for the members of the **[VisRowIndices](visrowindices-enumeration-visio.md)** enumeration. For a list of section index values, see the **AddSection** method or view the Visio type library for the members of the **[VisSectionIndices](vissectionindices-enumeration-visio.md)** enumeration.


