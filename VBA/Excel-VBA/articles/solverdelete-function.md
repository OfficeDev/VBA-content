---
title: SolverDelete Function
keywords: vbaxl10.chm5205195
f1_keywords:
- vbaxl10.chm5205195
ms.prod: excel
ms.assetid: 08d285ef-7c11-2429-3d91-61c75c515c72
ms.date: 06/08/2017
---


# SolverDelete Function

Deletes an existing constraint. Equivalent to clicking  **Solver** in the **Data** | **Analysis** group and then clicking **Delete** in the **Solver Parameters** dialog box.


 **Note**  The Solver add-in is not enabled by default. Before you can use this function, the Solver add-in must be enabled and installed. For information about how to do that, see  [Using the Solver VBA Functions](using-the-solver-vba-functions.md). After the Solver add-in is installed, you must establish a reference to the Solver add-in. In the Visual Basic Editor, with a module active, click  **References** on the **Tools** menu, and then select **Solver** under **Available References**. If  **Solver** does not appear under **Available References**, click  **Browse**, and then open Solver.xlam in the \Program Files\Microsoft Office\Office14\Library\SOLVER subfolder.


 **SolverDelete( _CellRef_**,  **_Relation_**,  **_FormulaText_)**

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

## Example

This example loads the previously calculated Solver model stored on Sheet1, deletes one of the constraints, and then solves the model again.


```vb
Worksheets("Sheet1").Activate 
SolverLoad loadArea:=Range("A33:A38") 
SolverDelete cellRef:=Range("C4:E6"), _ 
 relation:=4 
SolverSolve userFinish:=False
```


