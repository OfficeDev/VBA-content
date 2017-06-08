---
title: Report.Hwnd Property (Access)
keywords: vbaac10.chm13719
f1_keywords:
- vbaac10.chm13719
ms.prod: access
api_name:
- Access.Report.Hwnd
ms.assetid: e2d045f4-57bf-8681-0e00-bb5fe287136d
ms.date: 06/08/2017
---


# Report.Hwnd Property (Access)

You can use the  **hWnd** property to determine the handle (a unique **Long Integer** value) assigned by Microsoft Windows to the current window. Read/write **Long**.


## Syntax

 _expression_. **Hwnd**

 _expression_ A variable that represents a **Report** object.


## Remarks

You can use this property in Visual Basic when making calls to Windows application programming interface (API) functions or other external routines that require the  **hWnd** property as an argument. Many Windows functions require the **hWnd** property value of the current window as one of the arguments.


 **Note**  Because the value of this property can change while a program is running, don't store the  **hWnd** property value in a public variable.


## Example

The following example uses the  **hWnd** property with the Windows API **IsZoomed** function to determine if a window is maximized.


```vb
' Enter on single line in Declarations section of Module window. 
Declare Function IsZoomed Lib "user32" (ByVal hWnd As Long) As Long 
 
Sub Form_Activate() 
 Dim intWindowHandle As Long 
 intWindowHandle = Screen.ActiveForm.hWnd 
 If Not IsZoomed(intWindowHandle) Then 
 DoCmd.Maximize 
 End If 
End Sub
```


## See also


#### Concepts


[Report Object](report-object-access.md)

