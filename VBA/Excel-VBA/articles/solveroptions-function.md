---
title: SolverOptions Function
keywords: vbaxl10.chm5205227
f1_keywords:
- vbaxl10.chm5205227
ms.prod: excel
ms.assetid: 270d5440-ac1e-2436-b632-5877ede0820e
ms.date: 06/08/2017
---


# SolverOptions Function

Allows you to specify advanced options for your Solver model. This function and its arguments correspond to the options in the  **Solver Options** dialog box.


 **Note**  The Solver add-in is not enabled by default. Before you can use this function, you must have the Solver add-in enabled and installed. For information about how to do that, see  [Using the Solver VBA Functions](using-the-solver-vba-functions.md). After the Solver add-in is installed, you must establish a reference to the Solver add-in. In the Visual Basic Editor, with a module active, click  **References** on the **Tools** menu, and then select **Solver** under **Available References**. If  **Solver** does not appear under **Available References**, click  **Browse**, and then open Solver.xlam in the \Program Files\Microsoft Office\Office14\Library\SOLVER subfolder.


 **SolverOptions(** **MaxTime**,  **Iterations**,  **Precision**,  **AssumeLinear**,  **StepThru**,  **Estimates**,  **Derivatives**,  **SearchOption**,  **IntTolerance**,  **Scaling**,  **Convergence**,  **AssumeNonNeg**,  **PopulationSize**,  **RandomSeed**,  **MultiStart**,  **RequireBounds**,  **MutationRate**,  **MaxSubproblems**,  **MaxIntegerSols**,  **SolveWithout**,  **MaxTimeNoImp)**

 **MaxTime** Optional **Variant**. The maximum amount of time (in seconds) Solver will spend solving the problem. The value must be a positive integer. 
 
 **Iterations** Optional **Variant**. The maximum number of iterations Solver will use in solving the problem. The value must be a positive integer. 
 
 **Precision** Optional **Variant**. A number between 0 (zero) and 1 that specifies the degree of precisionwith which constraints (including integer constraints) must be satisfied. The default precision is 0.000001. A smaller number of decimal places (for example, 0.0001) indicates a lower degree of precision. In general, the higher the degree of precision you specify (the smaller the number), the more time Solver will take to reach solutions.
 
 **AssumeLinear** Optional **Variant**.  **True** to have Solver assume that the underlying model is linear. This speeds the solution process, but it should be used only if all the relationships in the model are linear. The default value is **False**.
 **StepThru** Optional **Variant**.  **True** to have Solver pause at each trial solution. You can pass Solver a macro to run at each pause by using the **_ShowRef_** argument of the **[SolverSolve](solversolve-function.md)** function. **False** to not have Solver pause at each trial solution. The default value is **False**.
 
 **Estimates** Optional **Variant**. Specifies the approach used to obtain initial estimates of the basic variables in each one-dimensional search: 1 represents tangent estimates, and 2 represents quadratic estimates. Tangent estimates use linear extrapolation from a tangent vector. Quadratic estimates use quadratic extrapolation; this may improve the results for highly nonlinear problems. The default value is 1 (tangent estimates).
 
 **Derivatives** Optional **Variant**. Specifies forward differencing or central differencing for estimates of partial derivatives of the objective and constraint functions: 1 represents forward differencing, and 2 represents central differencing. Central differencing requires more worksheet recalculations, but it may help with problems that generate a message saying that Solver could not improve the solution. With constraints whose values change rapidly near their limits, you should use central differencing. The default value is 1 (forward differencing).
 
 **SearchOption** Optional **Variant**. Use the  **_Search_** options to specify the search algorithm that will be used at each iteration to decide which direction to search in: 1 represents the Newton search method, and 2 represents the conjugate search method. Newton, which uses a quasi-Newton method, is the default search method.
 
 **IntTolerance** Optional **Variant**. A decimal number between 0 (zero) and 100 that specifies the  **Integer Optimality** percentage tolerance. This argument applies only if integer constraints have been defined; it specifies that Solver can stop if it has found a feasible integer solution whose objective is within this percentage of the best known bound on the objective of the true integer optimal solution. A larger percentage tolerance would tend to speed up the solution process.
 
 **Scaling** Optional **Variant**. If the objective or constraints differ by several orders of magnitude, for example, maximizing percentage of profit based on million-dollar investments, set this option  **True** to have Solver internally rescale the objective and constraint values to similar orders of magnitude during computation. If this option is **False**, Solver will perform its computations with the original values of the objective and constraints. The default value is  **True**.
 
 **Convergence** Optional **Variant**. A number between 0 (zero) and 1 that specifies the convergence tolerance for the  **GRG Nonlinear Solving** and **Evolutionary Solving** methods. For the GRG method, when the relative change in the target cell value is less than this tolerance for the last five iterations, Solver stops. For the Evolutionary method, when 99% or more of the members of the population have "fitness" values whose relative, that is percentage, difference is less than this tolerance, Solver stops. In both cases, Solver displays the message "Solver converged to the current solution. All constraints are satisfied."
 
 **AssumeNonNeg** Optional **Variant**.  **True** to have Solver assume a lower limit of 0 (zero) for all decision variable cells that do not have explicit lower limits in the **Constraint** list box (the cells must contain nonnegative values). **False** to have Solver use only the limits specified in the **Constraint** list box.
 
 **PopulationSize** Optional **Variant**.  **True** to have Solver assume a lower limit of 0 (zero) for all decision variable cells that do not have explicit lower limits in the **Constraint** list box (the cells must contain nonnegative values). **False** to have Solver use only the limits specified in the Constraint list box.
 
 **RandomSeed** Optional **Variant**. A positive integer specifies a fixed seed for the random number generator used by the  **Evolutionary Solving** method and the multistart method for global optimization. This means that Solver will find the same solution each time it is run on a model that has not changed. A zero value specifies that Solver should use a different seed for the random number generator each time it runs, which may yield different solutions each time it is run on a model that has not changed.
 
 **MultiStart** Optional **Variant**.  **True** to have Solver use multistart method for global optimization with the ** GRG Nonlinear Solving** method, when **[SolverSolve](solversolve-function.md)** is called. **False** to have Solver run the **GRG Solving** method only once, without multistart, when **[SolverSolve](solversolve-function.md)** is called.
 
 **RequireBounds** Optional **Variant**.  **True** to cause the Evolutionary Solving method and the multistart method to return immediately from a call to **[SolverSolve](solversolve-function.md)** with a value of 18 if any of the variables do not have both lower and upper bounds defined. **False** to have these methods attempt to solve the problem without bounds on all of the variables.
 
 **MutationRate** Optional **Variant**. A number between 0 (zero) and 1 that specifies the rate at which the  **Evolutionary Solving** method will make "mutations" to existing population members. A higher Mutation rate tends to increase the diversity of the population, and may yield better solutions.
 
 **MaxSubproblems** Optional **Variant**. The maximum number of subproblems Solver will explore in problems with integer constraints, and problems solved via the  **Evolutionary Solving** method. The value must be a positive integer.
 
 **MaxIntegerSols** Optional **Variant**. The maximum number of feasible (or integer feasible) solutions Solver will consider in problems with integer constraints, and problems solved via the  **Evolutionary Solving** method. The value must be a positive integer.
 
 **SolveWithout** Optional **Variant**.  **True** to have Solver ignore any integer constraints and solve the "relaxation" of the problem. **False** to have Solver use the integer constraints in solving the problem.
 
 **MaxTimeNoImp ** Optional **Variant**. When the  **Evolutionary Solving** method is used, the maximum amount of time (in seconds) Solver will continue solving without finding significantly improved solutions to add to the population. The value must be a positive integer.

## Example

This example sets the  **Precision** option to .001.


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


