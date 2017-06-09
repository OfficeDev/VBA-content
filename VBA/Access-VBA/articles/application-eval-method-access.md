---
title: Application.Eval Method (Access)
keywords: vbaac10.chm12537
f1_keywords:
- vbaac10.chm12537
ms.prod: access
api_name:
- Access.Application.Eval
ms.assetid: d02d5278-1ff3-c405-d579-7a58f2e1ea68
ms.date: 06/08/2017
---


# Application.Eval Method (Access)

You can use the  **Eval** function to evaluate an expression that results in a text string or a numeric value.


## Syntax

 _expression_. **Eval**( ** _StringExpr_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _StringExpr_|Required|**String**|An expression that evaluates to an alphanumeric text string. For example,  _stringexpr_ can be a function that returns a string or a numeric value. Or it can be a reference to a control on a form. The _stringexpr_ argument must evaluate to a string or numeric value; it can't evaluate to a Microsoft Access object.|

### Return Value

Variant


## Remarks

You can construct a string and then pass it to the  **Eval** function as if the string were an actual expression. The **Eval** function evaluates the string expression and returns its value. For example, `Eval("1 + 1")` returns 2.

If you pass to the  **Eval** function a string that contains the name of a function, the **Eval** function returns the return value of the function. For example, `Eval("Chr$(65)")` returns "A".


 **Note**  If you are passing the name of a function to the  **Eval** function, you must include parentheses after the name of the function in the _stringexpr_ argument. For example:




```vb
' ShowNames is user-defined function. 
Debug.Print Eval("ShowNames()") 
```




```vb
Debug.Print Eval("StrComp(""Joe"",""joe"", 1)")
```




```vb
Debug.Print Eval("Date()")
```

You can use the  **Eval** function in a calculated control on a form or report, or in a macro or module. The **Eval** function returns a **Variant** that is either a string or a numeric type.

The argument  _stringexpr_ must be an expression that is stored in a string. If you pass to the **Eval** function a string that doesn't contain a numeric expression or a function name but only a simple text string, a run-time error occurs. For example, `Eval("Smith")` results in an error.

You can use the  **Eval** function to determine the value stored in the **Value** property of a control. The following example passes a string containing a full reference to a control to the **Eval** function. It then displays the current value of the control in a dialog box.




```vb
Dim ctl As Control 
Dim strCtl As String 
 
Set ctl = Forms!Employees!LastName 
strCtl = "Forms!Employees!LastName" 
MsgBox ("The current value of " &; ctl.Name &; " is " &; Eval(strCtl))
```

You can use the  **Eval** function to access expression operators that aren't ordinarily available in Visual Basic. For example, you can't use the SQL operators **Between...And** or **In** directly in your code, but you can use them in an expression passed to the **Eval** function.

The next example determines whether the value of a ShipRegion control on an Orders form is one of several specified state abbreviations. If the field contains one of the abbreviations,  `intState` will be **True** (?1). Note that you use single quotation marks (') to include a string within another string.




```vb
Dim intState As Integer 
intState = Eval("Forms!Orders!ShipRegion In " _ 
 &; "('AK', 'CA', 'ID', 'WA', 'MT', 'NM', 'OR')")
```


## Example

The following example assumes that you have a series of 50 functions defined as A1, A2, and so on. This example uses the  **Eval** function to call each function in the series.


```vb
Sub CallSeries() 
 
 Dim intI As Integer 
 
 For intI = 1 To 50 
 Eval("A" &; intI &; "()") 
 Next intI 
 
End Sub
```

The next example triggers a Click event as if the user had clicked a button on a form. If the value of the button's  **OnClick** property begins with an equal sign (=), signifying that it s the name of a function, the **Eval** function calls the function, which is equivalent to triggering the **Click** event. If the value doesn't begin with an equal sign, then the value must name a macro. The **RunMacro** method of the **DoCmd** object runs the named macro.




```vb
Dim ctl As Control 
Dim varTemp As Variant 
 
Set ctl = Forms!Contacts!HelpButton 
If (Left(ctl.OnClick, 1) = "=") Then 
 varTemp = Eval(Mid(ctl.OnClick,2)) 
Else 
 DoCmd.RunMacro ctl.OnClick 
End If
```


## See also


#### Concepts


[Application Object](application-object-access.md)

