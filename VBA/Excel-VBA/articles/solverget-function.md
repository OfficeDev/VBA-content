---
title: SolverGet Function
keywords: vbaxl10.chm5205209
f1_keywords:
- vbaxl10.chm5205209
ms.prod: excel
ms.assetid: 3daf519c-06be-b200-7615-926e30fd2474
ms.date: 06/08/2017
---


# SolverGet Function

Returns information about current settings for Solver. The settings are specified in the  **Solver Parameters** and **Solver Options** dialog boxes.


 **Note**  The Solver add-in is not enabled by default. Before you can use this function, you must have the Solver add-in enabled and installed. For information about how to do that, see  [Using the Solver VBA Functions](using-the-solver-vba-functions.md). After the Solver add-in is installed, you must establish a reference to the Solver add-in. In the Visual Basic Editor, with a module active, click  **References** on the **Tools** menu, and then select **Solver** under **Available References**. If  **Solver** does not appear under **Available References**, click  **Browse**, and then open Solver.xlam in the \Program Files\Microsoft Office\Office14\Library\SOLVER subfolder.


 **SolverGet**( **_TypeNum_**,  **_SheetName_**)

 **TypeNum** Required **Integer**. A number specifying the type of information you want. The following settings are specified in the  **Solver Parameters** dialog box.


| <strong>TypeNum</strong> | <strong>Returns</strong>                                                                                                                                                                              |
|:-------------------------|:------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| 1                        | The reference in the  <strong>Set Target Cell</strong> box, or the #N/A error value if Solver has not been used on the active sheet.                                                                  |
| 2                        | A number corresponding to the  <strong>Equal To</strong> option: 1 represents Max, 2 represents Min, and 3 represents Value Of.                                                                       |
| 3                        | The value in the  <strong>Value Of</strong> box.                                                                                                                                                      |
| 4                        | The reference (as a multiple reference, if necessary) in the  <strong>By Changing Cells</strong> box.                                                                                                 |
| 5                        | The number of constraints.                                                                                                                                                                            |
| 6                        | An array of the left sides of the constraints, in text form.                                                                                                                                          |
| 7                        | An array of numbers corresponding to the relationships between the left and right sides of the constraints: 1 represents <=, 2 represents =, 3 represents >=, 4 represents int, and 5 represents bin. |
| 8                        | An array of the right sides of the constraints, in text form.                                                                                                                                         |
| 13                       | <strong>True</strong> if the <strong>Simple LP Solving</strong> method is selected; <strong>False</strong> if another Solving method is selected.                                                     |
| 20                       | <strong>True</strong> if the ** Make Unconstrained Variables Non-Negative** check box is selected; <strong>False</strong> if it is cleared.                                                           |

The following settings are specified in the  **Solver Options** dialog box.


| <strong>TypeNum</strong> | <strong>Returns</strong>                                                                                                                                                           |
|:-------------------------|:-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| 9                        | The  <strong>Max Time (Seconds)</strong> option (All Methods tab).                                                                                                                 |
| 10                       | The  <strong>Iterations</strong> option (All Methods tab).                                                                                                                         |
| 11                       | The  <strong>Constraint Precision</strong> option (All Methods tab).                                                                                                               |
| 12                       | The  <strong>Integer Optimality (%)</strong> option (All Methods tab).                                                                                                             |
| 14                       | <strong>True</strong> if the <strong>Show Iteration Results</strong> check box is selected; <strong>False</strong> if it is cleared.                                               |
| 15                       | <strong>True</strong> if the <strong>Use Automatic Scaling</strong> check box is selected; <strong>False</strong> if it is cleared (All Methods tab).                              |
| 16                       | A number corresponding to the type of estimates: 1 represents Tangent, and 2 represents Quadratic.                                                                                 |
| 17                       | A number corresponding to the  <strong>Derivatives</strong> option in the GRG Nonlinear tab: 1 represents Forward, and 2 represents Central (GRG Nonlinear tab).                   |
| 18                       | A number corresponding to the type of search: 1 represents Newton, and 2 represents Conjugate.                                                                                     |
| 19                       | The  <strong>Convergence</strong> tolerance (GRG Nonlinear tab and Evolutionary tab).                                                                                              |
| 21                       | The  <strong>Population Size</strong> option (GRG Nonlinear tab and Evolutionary tab).                                                                                             |
| 22                       | The  <strong>Random Seed</strong> option(GRG Nonlinear tab and Evolutionary tab).                                                                                                  |
| 23                       | <strong>True</strong> if the Use ** Multistart** check box is selected; <strong>False</strong> if it is cleared (GRG Nonlinear tab).                                               |
| 24                       | <strong>True</strong> if the <strong>Require Bounds on Variables</strong> check box is selected; <strong>False</strong> if it is cleared (GRG Nonlinear tab and Evolutionary tab). |
| 25                       | The  <strong>Mutation Rate</strong> option (Evolutionary tab).                                                                                                                     |
| 26                       | The  <strong>Max Subproblems</strong> option (All Methods tab).                                                                                                                    |
| 27                       | The  <strong>Max Feasible Solutions</strong> option (All Methods tab).                                                                                                             |
| 28                       | The  <strong>Ignore Integer Constraints</strong> option (All Methods tab).                                                                                                         |
| 29                       | The  <strong>Maximum Time without Improvement</strong> option (Evolutionary tab).                                                                                                  |

 **SheetName** Optional **Variant**. The name of the sheet that contains the Solver model for which you want information. If  **_SheetName_** is omitted, this sheet is assumed to be the active sheet.

## Example

This example displays a message if you have not used Solver on Sheet1.


```vb
Worksheets("Sheet1").Activate 
state = SolverGet(TypeNum:=1) 
If IsError(State) Then 
 MsgBox "You have not used Solver on the active sheet" 
End If
```


