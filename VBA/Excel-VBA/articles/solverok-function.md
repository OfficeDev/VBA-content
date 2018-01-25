---
title: SolverOk Function
keywords: vbaxl10.chm5205218
f1_keywords:
- vbaxl10.chm5205218
ms.prod: excel
ms.assetid: d24a6a7b-e4d9-e315-d0a6-8b7c80a38ede
ms.date: 06/08/2017
---

# SolverOk Function

Defines a basic Solver model. Equivalent to clicking  **Solver** in the **Data** | **Analysis** group and then specifying options in the **Solver Parameters** dialog box.


 **Note**  The Solver add-in is not enabled by default. Before you can use this function, you must have the Solver add-in enabled and installed. For information about how to do that, see  [Using the Solver VBA Functions](using-the-solver-vba-functions.md). After the Solver add-in is installed, you must establish a reference to the Solver add-in. In the Visual Basic Editor, with a module active, click  **References** on the **Tools** menu, and then select **Solver** under **Available References**. If  **Solver** does not appear under **Available References**, click  **Browse**, and then open Solver.xlam in the \Program Files\Microsoft Office\Office14\Library\SOLVER subfolder.


 **SolverOk** **( _SetCell_**,  **_MaxMinVal_**,  **_ValueOf_**,  **_ByChange_**, **_Engine_**,  **_EngineDesc_)**

 **SetCell** Optional **Variant**. Refers to a single cell on the active worksheet. Corresponds to the  **Set Target Cell** box in the **Solver Parameters** dialog box.
 **MaxMinVal** Optional **Variant**. Corresponds to the  **Max**,  **Min**, and  **Value** options in the **Solver Parameters** dialog box.


|**MaxMinVal**|**Specifies**|
|:-----|:-----|
|1|Maximize|
|2|Minimize|
|3|Match a specific value|
 **ValueOf** Optional **Variant**. If  **_MaxMinVal_** is 3, you must specify the value to which the target cell is matched.
 
 **ByChange** Optional **Variant**. The cell or range of cells that will be changed so that you will obtain the desired result in the target cell. Corresponds to the  **By Changing Cells** box in the **Solver Parameters** dialog box.
 
 **Engine** Optional **Variant**. The Solving method that should be used to solve the problem: 2 for the Simplex LP method, 1 for the GRG Nonlinear method, or 3 for the Evolutionary method. Corresponds to the  **Select a Solving Method** dropdown list in the **Solver Parameters** dialog box.
 
 **EngineDesc** Optional **Variant**. An alternate way to specify the Solving method that should be used to solve the problem as a string: "Simplex LP", "GRG Nonlinear", or "Evolutionary". Corresponds to the  **Select a Solving Method** dropdown list in the **Solver Parameters** dialog box.

## Example

This example uses the Solver functions to maximize gross profit in a business problem. The  **SolverOK** function defines a problem by specifying the **_SetCell_,  _MaxMinVal_**, and  **_ByChange_** arguments.


```vb
Worksheets("Sheet1").Activate 
SolverReset 
SolverOptions precision:=0.001 
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


