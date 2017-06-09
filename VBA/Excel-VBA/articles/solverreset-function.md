---
title: SolverReset Function
keywords: vbaxl10.chm5205232
f1_keywords:
- vbaxl10.chm5205232
ms.prod: excel
ms.assetid: 5c8f99e7-9451-3e72-1d93-4fcd72fc3e71
ms.date: 06/08/2017
---


# SolverReset Function

Resets all cell selections and constraints in the  **Solver Parameters** dialog box and restores all the settings in the **Solver Options** dialog box to their defaults. Equivalent to clicking **Reset All** in the **Solver Parameters** dialog box. The **SolverReset** function is called automatically when you call the **[SolverLoad](solverload-function.md)** function, if the **_Merge_** argument is **False** or omitted.


 **Note**  The Solver add-in is not enabled by default. Before you can use this function, you must have the Solver add-in enabled and installed. For information about how to do that, see  [Using the Solver VBA Functions](using-the-solver-vba-functions.md). After the Solver add-in is installed, you must establish a reference to the Solver add-in. In the Visual Basic Editor, with a module active, click  **References** on the **Tools** menu, and then select **Solver** under **Available References**. If  **Solver** does not appear under **Available References**, click  **Browse**, and then open Solver.xlam in the \Program Files\Microsoft Office\Office14\Library\SOLVER subfolder.


 **SolverReset**( )


## Example

This example resets the Solver settings to their defaults before defining a new problem.


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


