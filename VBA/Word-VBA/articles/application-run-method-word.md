---
title: Application.Run Method (Word)
keywords: vbawd10.chm158335421
f1_keywords:
- vbawd10.chm158335421
ms.prod: word
api_name:
- Word.Application.Run
ms.assetid: d7d89a15-caea-d055-c60a-0e31acdfc2aa
ms.date: 06/08/2017
---


# Application.Run Method (Word)

Runs a Visual Basic macro.


## Syntax

 _expression_ . **Run**( **_MacroName_** , **_varg1_** , **_varg2_** , **_varg3_** , **_varg4_** , **_varg5_** , **_varg6_** , **_varg7_** , **_varg8_** , **_varg9_** , **_varg10_** , **_varg11_** , **_varg12_** , **_varg13_** , **_varg14_** , **_varg15_** , **_varg16_** , **_varg17_** , **_varg18_** , **_varg19_** , **_varg20_** , **_varg21_** , **_varg22_** , **_varg23_** , **_varg24_** , **_varg25_** , **_varg26_** , **_varg27_** , **_varg28_** , **_varg29_** , **_varg30_** )

 _expression_ Required. A variable that represents an **[Application](application-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _MacroName_|Required| **String**|The name of the macro.|
| _varg1...varg30_|Optional| **Variant**|Macro parameter values. You can pass up to 30 parameter values to the specified macro.|

## Remarks

The MacroName parameter can be any combination of template, module, and macro name. For example, the following statements are all valid.


```vb
Application.Run "Normal.Module1.MAIN" 
Application.Run "MyProject.MyModule.MyProcedure" 
Application.Run "'My Document.doc'!ThisModule.ThisProcedure"
```

If you specify the document name, your code can only run macros in documents related to the current context ? not just any macro in any document.

Although Visual Basic code can call a macro directly (without using the  **Run** method), this method is useful when the macro name is stored in a variable. (For more information, see the example for this topic). The following three statements are functionally equivalent. The first two statements require a reference to Normal.dot, the project in which the called macro resides; the third statement, which uses the **Run** method, does not require a reference to the Normal.dot project.




```vb
Normal.Module2.Macro1 
Call Normal.Module2.Macro1 
Application.Run MacroName:="Normal.Module2.Macro1"
```


## Example

This example prompts the user to enter a template name, module name, macro name, and parameter value, and then it runs that macro.


```vb
Dim strTemplate As String 
Dim strModule As String 
Dim strMacro As String 
Dim strParameter As String 
 
strTemplate = InputBox("Enter the template name") 
strModule = InputBox("Enter the module name") 
strMacro = InputBox("Enter the macro name") 
strParameter = InputBox("Enter a parameter value") 
Application.Run MacroName:=strTemplate &; "." _ 
 &; strModule &; "." &; strMacro, _ 
 varg1:=strParameter
```


## See also


#### Concepts


[Application Object](application-object-word.md)

