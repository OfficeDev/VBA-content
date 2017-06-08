---
title: Using the Solver VBA Functions
ms.prod: excel
ms.assetid: 37d0aa49-2e5c-5efe-1c69-b5168af1f231
ms.date: 06/08/2017
---


# Using the Solver VBA Functions

Before you can use the Solver VBA functions from VBA, you must enable the Solver add-in in the  **Excel Options** dialog box.


1. Click the  **File** tab, and then click **Options** below the **Excel** tab.
    
2. In the  **Excel Options** dialog box, click **Add-Ins**.
    
3. In the  **Manage** drop-down box, select **Excel Add-ins**, and then click  **Go**.
    
4. In the  **Add-Ins** dialog box, select **Solver Add-in**, and then click OK.
    

After you have enabled the Solver add-in, Excel will auto-install the Add-in if it is not already installed, and the  **Solver** command will be added to the **Analysis** group on the **Data** tab in the ribbon.

Before you can use the Solver VBA functions in the Visual Basic Editor, you must establish a reference to the Solver add-in. In the Visual Basic Editor, with a module active, click  **References** on the **Tools** menu, and then select **Solver** under **Available References**. If  **Solver** does not appear under **Available References**, click  **Browse**, and then open Solver.xlam in the \Program Files\Microsoft Office\Office14\Library\SOLVER subfolder.
The following functions can be used to control the Solver add-in from VBA. Each function corresponds to an action that you can perform interactively, through the  **Solver Parameters**,  **Solver Options**, and  **Solver Results** dialog boxes of the Solver add-in.

-  [SolverAdd Function](solveradd-function.md)
    
-  [SolverChange Function](solverchange-function.md)
    
-  [SolverDelete Function](solverdelete-function.md)
    
-  [SolverFinish Function](solverfinish-function.md)
    
-  [SolverFinishDialog Function](solverfinishdialog-function.md)
    
-  [SolverGet Function](solverget-function.md)
    
-  [SolverLoad Function](solverload-function.md)
    
-  [SolverOk Function](solverok-function.md)
    
-  [SolverOkDialog Function](solverokdialog-function.md)
    
-  [SolverOptions Function](solveroptions-function.md)
    
-  [SolverReset Function](solverreset-function.md)
    
-  [SolverSave Function](solversave-function.md)
    
-  [SolverSolve Function](solversolve-function.md)
    

