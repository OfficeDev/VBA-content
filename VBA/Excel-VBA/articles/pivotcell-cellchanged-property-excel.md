---
title: PivotCell.CellChanged Property (Excel)
keywords: vbaxl10.chm692088
f1_keywords:
- vbaxl10.chm692088
ms.prod: excel
api_name:
- Excel.PivotCell.CellChanged
ms.assetid: fc2ba7b5-4dfe-4b05-922e-2ac079c5afb1
ms.date: 06/08/2017
---


# PivotCell.CellChanged Property (Excel)

Returns whether a PivotTable value cell has been edited or recalculated since the PivotTable report was created or the last commit operation was performed. Read-only


## Syntax

 _expression_ . **CellChanged**

 _expression_ A variable that represents a **[PivotCell](pivotcell-object-excel.md)** object.


### Return Value

 **[XlCellChangedState](xlcellchangedstate-enumeration-excel.md)**


## Remarks

The value of the  **CellChanged** property is **xlCellNotChanged** by default.

For PivotTable reports with non-OLAP data sources, the value of this property can be only either  **xlCellNotChanged** or **xlCellChanged** . It is **xlCellNotChanged** for cells that have not been edited, and is **xlCellChanged** for cells that have been edited. Discarding changes sets its value to **xlCellNotChanged** .

Applying and saving changes applies only to PivotTable reports with OLAP data sources. The following list of descriptions of the possible states of the  **CellChange** property apply only to PivotTables with OLAP data sources:


-  **xlCellNotChanged** - the cell has not been edited or recalculated (if the cell contains a formula) since the PivotTable was created, or since the last save or discard changes operation was performed.
    
-  **xlCellChanged** - the cell has been edited or recalculated since the PivotTable was created, or since the last apply changes or save changes operation was performed, but that change has not yet been applied (the **UPDATE CUBE** statement has not been run for it).
    
-  **xlCellChangeApplied** - the cell has been edited or recalculated since the PivotTable was created, or since the last apply changes, save changes, or discard changes operation was performed, and that change has been applied (the **UPDATE CUBE** statement has been run for it).
    
The following table lists descriptions of how different actions by the user affect the setting of the  **CellChanged** property in a PivotTable with an OLAP data source.



|**User Action**|**Setting of CellChanged property for cells without formulas**|**Setting of CellChanged property for cells with formulas**|
|:-----|:-----|:-----|
|Types a value or formula into a cell or multiple cells.|Set to  **xlCellChanged** for those cells.|Set to  **xlCellChanged** for those cells.|
|Recalculates a cell or multiple cells with a formula, either manually (F9) or automatically by Excel.|N/A|Set to  **xlCellChanged** for those cells.|
|Saves (commits) changes.|Set to  **xlCellNotChanged** for all edited cells without a formula.|Set to  **xlCellChangeApplied** for all edited cells with a formula.|
|Discards all changes.|Set to  **xlCellNotChanged** for all edited cells without a formula.|Set to  **xlCellNotChanged** for all edited cells with a formula.|
|Discards change in single cell.|Set to  **xlCellNotChanged** for that cell only.|Set to  **xlCellNotChanged** for that cell only.|
|Clears multiple cells in one operation.|Set to  **xlCellNotChanged** for all those cells.|Set to  **xlCellNotChanged** for all those cells.|
|Clears a cell.|Set to  **xlCellNotChanged** for that cell only.|Set to  **xlCellNotChanged** for that cell only.|
|Performs an undo operation that changes the value back to a previously edited value, before that value was applied.|Set to  **xlCellChanged** for all edited cells without a formula.|Set to  **xlCellChanged** for all edited cells with a formula.|
|Performs an undo operation that changes the value back to a previously edited value, after that value was applied.|Set to  **xlCellChangedApplied** for all edited cells without a formula.|Set to  **xlCellChangeApplied** for all edited cells with a formula.|

## See also


#### Concepts


[PivotCell Object](pivotcell-object-excel.md)

