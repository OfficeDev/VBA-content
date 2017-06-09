---
title: SolverOkDialog Function
keywords: vbaxl10.chm5205223
f1_keywords:
- vbaxl10.chm5205223
ms.prod: excel
ms.assetid: b16cad05-2213-ab11-ee9f-c3e09fe7f973
ms.date: 06/08/2017
---


# SolverOkDialog Function

Same as the  **SolverOK** function, but also displays the **Solver** dialog box.


 **Note**  The Solver add-in is not enabled by default. Before you can use this function, you must have the Solver add-in enabled and installed. For information about how to do that, see  [Using the Solver VBA Functions](using-the-solver-vba-functions.md). After the Solver add-in is installed, you must establish a reference to the Solver add-in. In the Visual Basic Editor, with a module active, click  **References** on the **Tools** menu, and then select **Solver** under **Available References**. If  **Solver** does not appear under **Available References**, click  **Browse**, and then open Solver.xlam in the \Program Files\Microsoft Office\Office14\Library\SOLVER subfolder.


 **SolverOkDialog** **( _SetCell_**,  **_MaxMinVal_**,  **_ValueOf_**,  **_ByChange_**,  **_Engine_**,  **_EngineDesc_)**

 **SetCell** Optional **Variant**. Refers to a single cell on the active worksheet. Corresponds to the  **Set Target Cell** box in the **Solver Parameters** dialog box.
 **MaxMinVal** Optional **Variant**. Corresponds to the  **Max**,  **Min**, and  **Value** options in the **Solver Parameters** dialog box.


|**MaxMinVal**|**Specifies**|
|:-----|:-----|
|1|Maximize|
|2|Minimize|
|3|Match a specific value|
 **ValueOf** Optional **Variant**. If  **_MaxMinVal_** is 3, you must specify the value that the target cell is matched to.
 **ByChange** Optional **Variant**. The cell or range of cells that will be changed so that you will obtain the desired result in the target cell. Corresponds to the  **By Changing Cells** box in the **Solver Parameters** dialog box.
 **Engine** Optional **Variant**. The Solving method that should be used to solve the problem: 1 for the Simplex LP method, 2 for the GRG Nonlinear method, or 3 for the Evolutionary method. Corresponds to the  **Select a Solving Method** dropdown list in the **Solver Parameters** dialog box.
 **ByChange** Optional **Variant**. An alternate way to specify the Solving method that should be used to solve the problem as a string: "Simplex LP", "GRG Nonlinear", or "Evolutionary". Corresponds to the  **Select a Solving Method** dropdown list in the **Solver Parameters** dialog box.

## Example

This example loads the previously calculated Solver model stored on Sheet1, resets all Solver options, and then displays the  **Solver Parameters** dialog box. From this point on, you can use Solver manually.


```vb
Worksheets("Sheet1").Activate 
SolverLoad LoadArea:=Range("A33:A38") 
SolverReset 
SolverOKDialog SetCell:=Range("TotalProfit") 
SolverSolve UserFinish:=False
```


