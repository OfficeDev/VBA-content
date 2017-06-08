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


|**TypeNum**|**Returns**|
|:-----|:-----|
|1|The reference in the  **Set Target Cell** box, or the #N/A error value if Solver has not been used on the active sheet.|
|2|A number corresponding to the  **Equal To** option: 1 represents Max, 2 represents Min, and 3 represents Value Of.|
|3|The value in the  **Value Of** box.|
|4|The reference (as a multiple reference, if necessary) in the  **By Changing Cells** box.|
|5|The number of constraints.|
|6|An array of the left sides of the constraints, in text form.|
|7|An array of numbers corresponding to the relationships between the left and right sides of the constraints: 1 represents <=, 2 represents =, 3 represents >=, 4 represents int, and 5 represents bin.|
|8|An array of the right sides of the constraints, in text form.|
|13| **True** if the **Simple LP Solving** method is selected; **False** if another Solving method is selected.|
|20| **True** if the ** Make Unconstrained Variables Non-Negative** check box is selected; **False** if it is cleared.|
The following settings are specified in the  **Solver Options** dialog box.


|**TypeNum**|**Returns**|
|:-----|:-----|
|9|The  **Max Time (Seconds)** option (All Methods tab).|
|10|The  **Iterations** option (All Methods tab).|
|11|The  **Constraint Precision** option (All Methods tab).|
|12|The  **Integer Optimality (%)** option (All Methods tab).|
|14| **True** if the **Show Iteration Results** check box is selected; **False** if it is cleared.|
|15| **True** if the **Use Automatic Scaling** check box is selected; **False** if it is cleared (All Methods tab).|
|16|A number corresponding to the type of estimates: 1 represents Tangent, and 2 represents Quadratic.|
|17|A number corresponding to the  **Derivatives** option in the GRG Nonlinear tab: 1 represents Forward, and 2 represents Central (GRG Nonlinear tab).|
|18|A number corresponding to the type of search: 1 represents Newton, and 2 represents Conjugate.|
|19|The  **Convergence** tolerance (GRG Nonlinear tab and Evolutionary tab).|
|21|The  **Population Size** option (GRG Nonlinear tab and Evolutionary tab).|
|22|The  **Random Seed** option(GRG Nonlinear tab and Evolutionary tab).|
|23| **True** if the Use ** Multistart** check box is selected; **False** if it is cleared (GRG Nonlinear tab).|
|24| **True** if the **Require Bounds on Variables** check box is selected; **False** if it is cleared (GRG Nonlinear tab and Evolutionary tab).|
|25|The  **Mutation Rate** option (Evolutionary tab).|
|26|The  **Max Subproblems** option (All Methods tab).|
|27|The  **Max Feasible Solutions** option (All Methods tab).|
|28|The  **Ignore Integer Constraints** option (All Methods tab).|
|29|The  **Maximum Time without Improvement** option (Evolutionary tab).|
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


