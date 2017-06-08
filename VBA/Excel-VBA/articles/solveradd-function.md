---
title: SolverAdd Function
keywords: vbaxl10.chm5205183
f1_keywords:
- vbaxl10.chm5205183
ms.prod: excel
ms.assetid: c20e0a78-113e-254f-428f-0dc1bdc817c2
ms.date: 06/08/2017
---


# SolverAdd Function

Adds a constraint to the current problem. Equivalent to clicking  **Solver** in the **Data** | **Analysis** group and then clicking **Add** in the **Solver Parameters** dialog box.


 **Note**  The Solver add-in is not enabled by default. Before you can use this function, the Solver add-in must be enabled and installed. For information about how to do that, see  [Using the Solver VBA Functions](using-the-solver-vba-functions.md). After the Solver add-in is installed, you must establish a reference to the Solver add-in. In the Visual Basic Editor, with a module active, click  **References** on the **Tools** menu, and then select **Solver** under **Available References**. If  **Solver** does not appear under **Available References**, click  **Browse**, and then open Solver.xlam in the \Program Files\Microsoft Office\Office14\Library\SOLVER subfolder.


 **SolverAdd( _CellRef_**,  **_Relation_**,  **_FormulaText_)**

 **CellRef** Required **Variant**. A reference to a cell or a range of cells that forms the left side of a constraint.
 **Relation** Required **Integer**. The arithmetic relationship between the left and right sides of the constraint. If you choose 4, 5, or 6,  **_CellRef_** must refer to decision variable cells, and **_FormulaText_** should not be specified.


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

After constraints are added, you can manipulate them with the  **[SolverChange](solverchange-function.md)** and  **[SolverDelete](solverdelete-function.md)** functions.


## Example

This example uses the Solver functions to maximize gross profit in a business problem. The  **SolverAdd** function is used to add three constraints to the current problem.


```vb
Worksheets("Sheet1").Activate 
SolverReset 
SolverOptions precision:=0.001 
SolverOK setCell:=Range("TotalProfit"), _ 
 maxMinVal:=1, _ 
 byChange:=Range("C4:E6") 
SolverAdd cellRef:=Range("F4:F6"), _ 
 relation:=1, _ 
 formulaText:=100 
SolverAdd cellRef:=Range("C4:E6"), _ 
 relation:=3, _ 
 formulaText:=0 
SolverAdd cellRef:=Range("C4:E6"), _ 
 relation:=4 
SolverSolve userFinish:=False 
SolverSave saveArea:=Range("A33")
```


