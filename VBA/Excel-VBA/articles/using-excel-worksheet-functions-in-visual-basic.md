---
title: Using Excel Worksheet Functions in Visual Basic
keywords: vbaxl10.chm81924
f1_keywords:
- vbaxl10.chm81924
ms.prod: excel
ms.assetid: 46e6ba32-8a58-509c-03e8-a23c41b0a400
ms.date: 06/08/2017
---


# Using Excel Worksheet Functions in Visual Basic

You can use most Microsoft Excel worksheet functions in your Visual Basic statements. For a list of the worksheet functions you can use, see  [List of Worksheet Functions Available to Visual Basic](list-of-worksheet-functions-available-to-visual-basic.md).


 **Note**  Some worksheet functions are not useful in Visual Basic. For example, the  **Concatenate** function is not needed because in Visual Basic you can use the **&;** operator to join multiple text values.


## Calling a Worksheet Function from Visual Basic

In Visual Basic, the Excel worksheet functions are available through the  **WorksheetFunction** object.

The following  **Sub** procedure uses the **Min** worksheet function to determine the smallest value in a range of cells. First, the variable `myRange` is declared as a **Range** object, and then it is set to range A1:C10 on Sheet1. Another variable, `answer`, is assigned the result of applying the  **Min** function to `myRange`. Finally, the value of  `answer` is displayed in a message box.




```vb
Sub UseFunction() 
 Dim myRange As Range 
 Set myRange = Worksheets("Sheet1").Range("A1:C10") 
 answer = Application.WorksheetFunction.Min(myRange) 
 MsgBox answer 
End Sub
```

If you use a worksheet function that requires a range reference as an argument, you must specify a  **Range** object. For example, you can use the **Match** worksheet function to search a range of cells. In a worksheet cell, you would enter a formula such as =MATCH(9,A1:A10,0). However, in a Visual Basic procedure, you would specify a **Range** object to get the same result.




```vb
Sub FindFirst() 
 myVar = Application.WorksheetFunction _ 
 .Match(9, Worksheets(1).Range("A1:A10"), 0) 
 MsgBox myVar 
End Sub
```


 **Note**  Visual Basic functions do not use the  **WorksheetFunction** qualifier. A function may have the same name as a Microsoft Excel function and yet work differently. For example, `Application.WorksheetFunction.Log` and `Log` will return different values.


## Inserting a Worksheet Function into a Cell

To insert a worksheet function into a cell, you specify the function as the value of the  **Formula** property of the corresponding **Range** object. In the following example, the RAND worksheet function (which generates a random number) is assigned to the **Formula** property of range A1:B3 on Sheet1 in the active workbook.


```vb
Sub InsertFormula() 
 Worksheets("Sheet1").Range("A1:B3").Formula = "=RAND()" 
End Sub
```


## Example

This example uses the worksheet function  **Pmt** to calculate a home mortgage loan payment. Notice that this example uses the **InputBox** method instead of the **InputBox** function so that the method can perform type checking. The **Static** statements cause Visual Basic to retain the values of the three variables; these are displayed as default values the next time you run the program.


```vb
Static loanAmt 
Static loanInt 
Static loanTerm 
loanAmt = Application.InputBox _ 
 (Prompt:="Loan amount (100,000 for example)", _ 
 Default:=loanAmt, Type:=1) 
loanInt = Application.InputBox _ 
 (Prompt:="Annual interest rate (8.75 for example)", _ 
 Default:=loanInt, Type:=1) 
loanTerm = Application.InputBox _ 
 (Prompt:="Term in years (30 for example)", _ 
 Default:=loanTerm, Type:=1) 
payment = Application.WorksheetFunction _ 
 .Pmt(loanInt / 1200, loanTerm * 12, loanAmt) 
MsgBox "Monthly payment is " &; Format(payment, "Currency")
```


