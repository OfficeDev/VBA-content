
# SolverLoad Function

 **Last modified:** July 28, 2015

Loads existing Solver model parameters that have been saved to the worksheet.

 **Note**  The Solver add-in is not enabled by default. Before you can use this function, you must have the Solver add-in enabled and installed. For information about how to do that, see  [Using the Solver VBA Functions](37d0aa49-2e5c-5efe-1c69-b5168af1f231.md). After the Solver add-in is installed, you must establish a reference to the Solver add-in. In the Visual Basic Editor, with a module active, click  **References** on the **Tools** menu, and then select **Solver** under **Available References**. If  **Solver** does not appear under **Available References**, click  **Browse**, and then open Solver.xlam in the \Program Files\Microsoft Office\Office14\Library\SOLVER subfolder.

 **SolverLoad( _LoadArea_**,  **_Merge_)**
 **LoadArea** Required **Variant**. A reference on the active worksheet to a range of cells from which you want to load a complete problem specification. The first cell in the  **_LoadArea_** contains a formula for the **Set Target Cell** box in the **Solver Parameters** dialog box; the second cell contains a formula for the **By Changing Cells** box; subsequent cells contain constraints in the form of logical formulas. The last cell optionally contains an array of Solver option values. For more information, see ** [SolverOptions](270d5440-ac1e-2436-b632-5877ede0820e.md)**. The range represented by the argument  **_LoadArea_** can be on any worksheet, but you must specify the worksheet if it is not the active sheet. For example, `SolverLoad("Sheet2!A1:A3")` loads a model from Sheet2 even if it is not the active sheet.
 **Merge** Optional **Variant**. A logical value corresponding to either the  **Merge** button or the **Replace** button in the dialog box that appears after you select the **LoadArea** reference and click **OK**. If  **True**, the variable cell selections and constraints from the LoadArea are merged with the currently defined variables and constraints. If  **False** or omitted, the current model specifications and options are erased (equivalent to a call to the ** [SolverReset](5c8f99e7-9451-3e72-1d93-4fcd72fc3e71.md)** function) before the new specifications are loaded.

## Example

This example loads the previously calculated Solver model stored on Sheet1, changes one of the constraints, and then solves the model again.


```
Worksheets("Sheet1").Activate 
SolverLoad loadArea:=Range("A33:A38") 
SolverChange cellRef:=Range("F4:F6"), _ 
 relation:=1, _ 
 formulaText:=200 
SolverSolve userFinish:=False
```

