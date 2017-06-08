---
title: SolverFinish Function
keywords: vbaxl10.chm5205200
f1_keywords:
- vbaxl10.chm5205200
ms.prod: excel
ms.assetid: 8657931a-a618-bc90-d850-6d37707e0a68
ms.date: 06/08/2017
---


# SolverFinish Function

Tells Microsoft Office Excel what to do with the results and what kind of report to create when the solution process is completed.


 **Note**  The Solver add-in is not enabled by default. Before you can use this function, you must have the Solver add-in enabled and installed. For information about how to do that, see  [Using the Solver VBA Functions](using-the-solver-vba-functions.md). After the Solver add-in is installed, you must establish a reference to the Solver add-in. In the Visual Basic Editor, with a module active, click  **References** on the **Tools** menu, and then select **Solver** under **Available References**. If  **Solver** does not appear under **Available References**, click  **Browse**, and then open Solver.xlam in the \Program Files\Microsoft Office\Office14\Library\SOLVER subfolder.


 **SolverFinish**( **_KeepFinal_**,  **_ReportArray_**, **_OutlineReports_**)

 **KeepFinal** Optional **Variant**. Can be either 1 or 2. If  **_KeepFinal_** is 1 or omitted, the final solution values are kept in the changing cells, replacing any former values. If **_KeepFinal_** is 2, the final solution values are discarded, and the former values are restored.
 **ReportArray** Optional **Variant**. The kind of report that Excel will create when Solver is finished:

- When the Simplex LP or GRG Nonlinear Solving method is used, 1 creates an Answer report, 2 creates a Sensitivity report, and 3 creates a Limit report. 
    
- When the Evolutionary Solving method is used, 1 creates an Answer report, and 2 creates a Population report.
    
- When  **[SolverSolve](solversolve-function.md)** returns 5 (Solver could not find a feasible solution), 1 creates a Feasibility Report, and 2 creates a Feasibility-Bounds report.
    
- When  **SolverSolve** returns 7 (the linearity conditions are not satisfied), 1 creates a Linearity report.
    
 Use the **Array** function to specify the reports you want to display — for example, `ReportArray:= Array(1,3)`.
 **OutlineReports** Optional **Variant**. Can be either  **True** or **False**. If  **_OutlineReports_** is **False** or omitted, reports are produced in the "regular" format, without outlining. If **_OutlineReports_** is **True**, reports are produced with outlined groups corresponding to the cell ranges you've entered for decision variables and constraints. 

## Example

This example loads the previously calculated Solver model stored on Sheet1, solves the model again, and then generates an answer report on a new worksheet.


```vb
Worksheets("Sheet1").Activate 
SolverLoad LoadArea:=Range("A33:A38") 
SolverSolve UserFinish:=True 
SolverFinish KeepFinal:=1, ReportArray:=Array(1)
```


