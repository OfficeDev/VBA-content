---
title: SolverSave Function
keywords: vbaxl10.chm5205236
f1_keywords:
- vbaxl10.chm5205236
ms.prod: excel
ms.assetid: 177dcfb7-b223-c172-d4d6-9cab534a8fa5
ms.date: 06/08/2017
---


# SolverSave Function

Saves the Solver problem specifications on the worksheet.


 **Note**  The Solver add-in is not enabled by default. Before you can use this function, you must have the Solver add-in enabled and installed. For information about how to do that, see  [Using the Solver VBA Functions](using-the-solver-vba-functions.md). After the Solver add-in is installed, you must establish a reference to the Solver add-in. In the Visual Basic Editor, with a module active, click  **References** on the **Tools** menu, and then select **Solver** under **Available References**. If  **Solver** does not appear under **Available References**, click  **Browse**, and then open Solver.xlam in the \Program Files\Microsoft Office\Office14\Library\SOLVER subfolder.


 **SolverSave( _SaveArea_)**

 **SaveArea** Required **Variant**. The range of cells where the Solver model is to be saved. If this is a single-cell range, Solver uses as many cells as it needs to save the model, in a column starting with the specified cell. If this is a multi-cell range, Solver uses only cells within that range, even if the model cannot be entirely saved, The range represented by the  **_SaveArea_** argument can be on any worksheet, but you must specify the worksheet if it is not the active sheet. For example, `SolverSave("Sheet2!A1:A3")` saves the model on Sheet2 even if Sheet2 is not the active sheet.

## Example

This example uses the Solver functions to maximize gross profit in a business problem. The  **SolverSave** function saves the current problem to a range on the active worksheet.


```vb
Worksheets("Sheet1").Activate 
SolverReset 
SolverOptions Precision:=0.001 
SolverOK SetCell:=Range("TotalProfit"), _ 
 MaxMinVal:=1, _ 
 ByChange:=Range("C4:E6") 
SolverAdd CellRef:=Range("F4:F6"), _ 
 Relation:=1, _ 
 FormulaText:=100 
SolverAdd CellRef:=Range("C4:E6"), _ 
 Relation:=3, _ 
 FormulaText:=0 
SolverAdd CellRef:=Range("C4:E6"), _ 
 Relation:=4 
SolverSolve UserFinish:=False 
SolverSave SaveArea:=Range("A33")
```


