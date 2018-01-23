---
title: Application.InputBox Method (Excel)
keywords: vbaxl10.chm133149
f1_keywords:
- vbaxl10.chm133149
ms.prod: excel
api_name:
- Excel.Application.InputBox
ms.assetid: d3bd2f3a-7fed-20fa-918d-a71e2a2a1d49
ms.date: 06/08/2017
---


# Application.InputBox Method (Excel)

Displays a dialog box for user input. Returns the information entered in the dialog box.


## Syntax

 _expression_. **InputBox** (**_Prompt_**, **_Title_**, **_Default_**, **_Left_**, **_Top_**, **_HelpFile_**, **_HelpContextID_**, **_Type_**)

 _expression_ A variable that represents an **Application** object.


### Parameters

|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Prompt_|Required| **String**|The message to be displayed in the dialog box. This can be a string, a number, a date, or a Boolean value (Microsoft Excel automatically coerces the value to a **String** before it is displayed).|
| _Title_|Optional| **Variant**|The title for the input box. If this argument is omitted, the default title is "Input."|
| _Default_|Optional| **Variant**|Specifies a value that will appear in the text box when the dialog box is initially displayed. If this argument is omitted, the text box is left empty. This value can be a **[Range](range-object-excel.md)** object.|
| _Left_|Optional| **Variant**|Specifies an *x* position for the dialog box in relation to the upper-left corner of the screen, in points.|
| _Top_|Optional| **Variant**|Specifies a *y* position for the dialog box in relation to the upper-left corner of the screen, in points.|
| _HelpFile_|Optional| **Variant**|The name of the Help file for this input box. If the _HelpFile_ and _HelpContextID_ arguments are present, a Help button will appear in the dialog box.|
| _HelpContextID_|Optional| **Variant**|The context ID number of the Help topic in _HelpFile_.|
| _Type_|Optional| **Variant**|Specifies the return data type. If this argument is omitted, the dialog box returns text.|

### Return value

Variant


## Remarks

The following table lists the values that can be passed in the Type argument. Can be one or a sum of the values. For example, for an input box that can accept both text and numbers, set _Type_ to 1 + 2.

|**Value**|**Meaning**|
|:-----|:-----|
|0|A formula|
|1|A number|
|2|Text (a string)|
|4|A logical value (**True** or **False**)|
|8|A cell reference, as a **Range** object|
|16|An error value, such as #N/A|
|64|An array of values|


Use **InputBox** to display a simple dialog box so that you can enter information to be used in a macro. The dialog box has an **OK** button and a **Cancel** button. If you select the **OK** button, **InputBox** returns the value entered in the dialog box. If you select the **Cancel** button, **InputBox** returns **False**.

If _Type_ is 0, **InputBox** returns the formula in the form of text; for example, `=2*PI()/360`. If there are any references in the formula, they are returned as A1-style references. (Use **[ConvertFormula](application-convertformula-method-excel.md)** to convert between reference styles.)

If _Type_ is 8, **InputBox** returns a **Range** object. You must use the **Set** statement to assign the result to a **Range** object, as shown in the following example.

```vb
Set myRange = Application.InputBox(prompt := "Sample", type := 8)
```

If you do not use the **Set** statement, the variable is set to the value in the range, rather than the **Range** object itself.

If you use the **InputBox** method to ask the user for a formula, you must use the **[FormulaLocal](range-formulalocal-property-excel.md)** property to assign the formula to a **Range** object. The input formula will be in the user's language.

The **InputBox** method differs from the **InputBox** function in that it allows selective validation of the user's input, and it can be used with Excel objects, error values, and formulas. Notice that `Application.InputBox` calls the **InputBox** method; `InputBox` with no object qualifier calls the **InputBox** function.


## Example

This example prompts the user for a number.

```
myNum = Application.InputBox("Enter a number")
```

This example prompts the user to select a cell on Sheet1. The example uses the _Type_ argument to ensure that the return value is a valid cell reference (a **Range** object).

```vb
Worksheets("Sheet1").Activate 
Set myCell = Application.InputBox( _ 
    prompt:="Select a cell", Type:=8)
```

**Sample code provided by:** Holy Macro! Books, [Holy Macro! It's 2,500 Excel VBA Examples](http://www.mrexcel.com/store/index.php?l=product_detail&p=1)

This example uses an InputBox for the user to select a range to pass to the user-defined function 'MyFunction', which multiplies three values in a range together and returns the result.

```vb
Sub Cbm_Value_Select()
   'Set up the variables.
   Dim rng As Range
   
   'Use the InputBox dialog to set the range for MyFunction, with some simple error handling.
   Set rng = Application.InputBox("Range:", Type:=8)
   If rng.Cells.Count <> 3 Then
     MsgBox "Length, width and height are needed -" &; _
         vbLf &; "please select three cells!"
      Exit Sub
   End If
   
   'Call MyFunction by value using the active cell.
   ActiveCell.Value = MyFunction(rng)
End Sub

Function MyFunction(rng As Range) As Double
   MyFunction = rng(1) * rng(2) * rng(3)
End Function
```

## About the contributor
<a name="AboutContributor"> </a>

*Holy Macro! Books* publishes entertaining books for people who use Office. See the complete catalog at MrExcel.com. 

## See also

#### Concepts

[Application Object](application-object-excel.md)

