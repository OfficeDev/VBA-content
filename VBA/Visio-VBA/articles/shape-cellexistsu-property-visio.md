---
title: Shape.CellExistsU Property (Visio)
keywords: vis_sdr.chm11251945
f1_keywords:
- vis_sdr.chm11251945
ms.prod: visio
api_name:
- Visio.Shape.CellExistsU
ms.assetid: da26e913-39c5-7af5-194d-3bb5dca76678
ms.date: 06/08/2017
---


# Shape.CellExistsU Property (Visio)

Determines whether a particular ShapeSheet cell exists in the scope of the search. Read-only.


## Syntax

 _expression_ . **CellExistsU**( **_localeIndependentCellName_** , **_fExistsLocally_** )

 _expression_ A variable that represents a **Shape** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _localeIndependentCellName_|Required| **String**|The universal name of the ShapeSheet cell for which you want to search.|
| _fExistsLocally_|Required| **Integer**|The scope of the search.|

### Return Value

Integer


## Remarks

The  _localeIndependentCellName_ argument must specify a universal cell name. To search for a cell by section, row, and column index, use the **CellsSRCExists** property.

The  _fExistsLocally_ argument specifies the scope of the search.




- If  _fExistsLocally_ is non-zero ( **True** ), the **CellExistsU** property value is **True** only if the object contains the cell locally; if the cell is inherited, the **CellExistsU** property value is **False** .
    
- If  _fExistsLocally_ is zero ( **False** ), the **CellExistsU** property value is **True** if the object either contains or inherits the cell.
    


For a list of cell index values, view the Visio type library for the members of class  **[VisCellIndices](viscellindices-enumeration-visio.md)** .




 **Note**  Beginning with Microsoft Visio 2000, you can use both local and universal names to refer to Visio shapes, masters, documents, pages, rows, add-ons, cells, hyperlinks, styles, fonts, master shortcuts, UI objects, and layers. When a user names a shape, for example, the user is specifying a local name. Beginning with Microsoft Office Visio 2003, the ShapeSheet spreadsheet displays only universal names in cell formulas and values. (In prior versions, universal names were not visible in the user interface.) 

As a developer, you can use universal names in a program when you don't want to change a name each time a solution is localized. Use the  **CellExists** property to determine if a cell exists by passing either the cell's local name or its universal name. Use the **CellExistsU** property to determine if a cell exists by passing the cell's universal name.


