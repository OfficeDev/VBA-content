---
title: SolverChange Function
keywords: vbaxl10.chm5205187
f1_keywords:
- vbaxl10.chm5205187
ms.prod: excel
ms.assetid: 773c68cc-5d37-b8ff-c895-61fca75da5c1
ms.date: 06/08/2017
---


# SolverChange Function

Changes an existing constraint. Equivalent to clicking  **Solver** in the **Data** | **Analysis** group and then clicking **Change** in the **Solver Parameters** dialog box.


 **Note**  The Solver add-in is not enabled by default. Before you can use this function, the Solver add-in must be enabled and installed. For information about how to do that, see  [Using the Solver VBA Functions](using-the-solver-vba-functions.md). After the Solver add-in is installed, you must establish a reference to the Solver add-in. In the Visual Basic Editor, with a module active, click  **References** on the **Tools** menu, and then select **Solver** under **Available References**. If  **Solver** does not appear under **Available References**, click  **Browse**, and then open Solver.xlam in the \Program Files\Microsoft Office\Office14\Library\SOLVER subfolder.


 **SolverChange( _CellRef_**,  **_Relation_**,  **_FormulaText_)**

 **CellRef** Required **Variant**. A reference to a cell or a range of cells that forms the left side of a constraint.
 **Relation** Required **Integer**. The arithmetic relationship between the left and right sides of the constraint. If you choose 4 or 5,  **_CellRef_** must refer to adjustable (changing) cells, and **_FormulaText_** should not be specified.


|**Relation**|**Arithmetic relationship**|
|:-----|:-----|
|1|<=|
|2|=|
|3|>=|
|4|Cells referenced by  **_CellRef_** must have final values that are integers.|
|5|Cells referenced by  **_CellRef_** must have final values of either 0 (zero) or 1.|
|6|Cells referenced by  **_CellRef_** must have final values that are all different and integers.|
 **FormulaText** Optional **Variant**. The right side of the constraint.

## Remarks

If  **_CellRef_** and **_Relation_** do not match an existing constraint, you must use the **[SolverDelete](solverdelete-function.md)** and  **[SolverAdd](solveradd-function.md)** functions to change the constraint.


## Example

This example loads the previously calculated Solver model stored on Sheet1, changes one of the constraints, and then solves the model again.


```vb
Worksheets("Sheet1").Activate 
SolverLoad loadArea:=Range("A33:A38") 
SolverChange cellRef:=Range("F4:F6"), _ 
 relation:=1, _ 
 formulaText:=200 
SolverSolve userFinish:=False
```


