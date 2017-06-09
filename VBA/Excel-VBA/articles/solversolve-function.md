---
title: SolverSolve Function
keywords: vbaxl10.chm5205240
f1_keywords:
- vbaxl10.chm5205240
ms.prod: excel
ms.assetid: 40ef53c8-ff54-bdc8-9f8b-bf9a4445ce51
ms.date: 06/08/2017
---


# SolverSolve Function

Begins a Solver solution run. Equivalent to clicking  **Solve** in the **Solver Parameters** dialog box.


 **Note**  The Solver add-in is not enabled by default. Before you can use this function, you must have the Solver add-in enabled and installed. For information about how to do that, see  [Using the Solver VBA Functions](using-the-solver-vba-functions.md). After the Solver add-in is installed, you must establish a reference to the Solver add-in. In the Visual Basic Editor, with a module active, click  **References** on the **Tools** menu, and then select **Solver** under **Available References**. If  **Solver** does not appear under **Available References**, click  **Browse**, and then open Solver.xlam in the \Program Files\Microsoft Office\Office14\Library\SOLVER subfolder.


 **SolverSolve( _UserFinish_**,  **_ShowRef_)**

 **UserFinish** Optional **Variant**.  **True** to return the results without displaying the **Solver Results** dialog box. **False** or omitted to return the results and display the **Solver Results** dialog box.
 **ShowRef** Optional **Variant**. You can pass the name of a macro (as a string) as the  **_ShowRef_** argument. This macro is then called, in lieu of displaying the **Show Trial Solution** dialog box, whenever Solver pauses for any of the reasons listed below.The **_ShowRef_** macro must have the signature **Function  _name_ (Reason As Integer)**. The argument  **Reason** is an integer value from 1 to 5:

1. Function called (on every iteration) because the  **Show Iteration Results** box in the **Solver Options** dialog box was checked, or function called because the user pressed ESC to interrupt the Solver.
    
2. Function called because the  **Max Time** limit in the **Solver Options** dialog box was exceeded.
    
3. Function called because the  **Iterations** limit in the **Solver Options** dialog box was exceeded.
    
4. Function called because the  **Maximum Subproblems** limit in the **Solver Options** dialog box was exceeded.
    
5. Function called because the  **Maximum Feasible Solutions** limit in the **Solver Options** dialog box was exceeded.
    
The macro function must return 1 if Solver should stop (same as the  **Stop** button in the **Show Trial Solution** dialog box), or 0 if Solver should continue running (same as the **Continue** button).The **_ShowRef_** macro can inspect the current solution values on the worksheet, or take other actions such as saving or charting the intermediate values. However, it should not alter the values in the variable cells, or alter the formulas in the objective and constraint cells, as this could adversely affect the solution process.

## SolverSolve Return Value

If a Solver problem has not been completely defined,  **SolverSolve** returns the #N/A error value. Otherwise the Solver runs, and **SolverSolve** returns an integer value corresponding to the message that appears in the **Solver Results** dialog box:



|0|Solver found a solution. All constraints and optimality conditions are satisfied.|
|1|Solver has converged to the current solution. All constraints are satisfied.|
|2|Solver cannot improve the current solution. All constraints are satisfied.|
|3|Stop chosen when the maximum iteration limit was reached.|
|4|The Objective Cell values do not converge.|
|5|Solver could not find a feasible solution.|
|6|Solver stopped at user's request.|
|7|The linearity conditions required by this LP Solver are not satisfied.|
|8|The problem is too large for Solver to handle.|
|9|Solver encountered an error value in a target or constraint cell.|
|10|Stop chosen when the maximum time limit was reached.|
|11|There is not enough memory available to solve the problem.|
|13|Error in model. Please verify that all cells and constraints are valid.|
|14|Solver found an integer solution within tolerance. All constraints are satisfied.|
|15|Stop chosen when the maximum number of feasible [integer] solutions was reached.|
|16|Stop chosen when the maximum number of feasible [integer] subproblems was reached.|
|17|Solver converged in probability to a global solution.|
|18|All variables must have both upper and lower bounds.|
|19|Variable bounds conflict in binary or alldifferent constraint.|
|20|Lower and upper bounds on variables allow no feasible solution.|

## Example

This example uses the Solver functions to maximize gross profit in a business problem. The  **SolverSolve** function begins the Solver solution run. Solver calls the function `ShowTrial` when any of the five conditions described above occurs; the function simply displays a message with the integer value 1 through 5.


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
SolverSolve UserFinish:=False, ShowRef:= "ShowTrial" 
SolverSave SaveArea:=Range("A33") 
 
Function ShowTrial(Reason As Integer) 
  Msgbox Reason 
  ShowTrial = 0 
End Function
```


