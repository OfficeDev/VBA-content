
# Application.Run Method (Word)

 **Last modified:** July 28, 2015

 **In this article**
 [Syntax](#sectionSection0)
 [Remarks](#sectionSection1)
 [Example](#sectionSection2)


Runs a Visual Basic macro.


## Syntax
<a name="sectionSection0"> </a>

 _expression_. **Run**( **_MacroName_**,  **_varg1_**,  **_varg2_**,  **_varg3_**,  **_varg4_**,  **_varg5_**,  **_varg6_**,  **_varg7_**,  **_varg8_**,  **_varg9_**,  **_varg10_**,  **_varg11_**,  **_varg12_**,  **_varg13_**,  **_varg14_**,  **_varg15_**,  **_varg16_**,  **_varg17_**,  **_varg18_**,  **_varg19_**,  **_varg20_**,  **_varg21_**,  **_varg22_**,  **_varg23_**,  **_varg24_**,  **_varg25_**,  **_varg26_**,  **_varg27_**,  **_varg28_**,  **_varg29_**,  **_varg30_**)

 _expression_Required. A variable that represents an  ** [Application](d1cf6f8f-4e88-bf01-93b4-90a83f79cb44.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|MacroName|Required| **String**|The name of the macro.|
|varg1...varg30|Optional| **Variant**|Macro parameter values. You can pass up to 30 parameter values to the specified macro.|

## Remarks
<a name="sectionSection1"> </a>

The MacroName parameter can be any combination of template, module, and macro name. For example, the following statements are all valid.


```
Application.Run "Normal.Module1.MAIN" 
Application.Run "MyProject.MyModule.MyProcedure" 
Application.Run "'My Document.doc'!ThisModule.ThisProcedure"
```

If you specify the document name, your code can only run macros in documents related to the current context â€” not just any macro in any document.

Although Visual Basic code can call a macro directly (without using the  **Run** method), this method is useful when the macro name is stored in a variable. (For more information, see the example for this topic). The following three statements are functionally equivalent. The first two statements require a reference to Normal.dot, the project in which the called macro resides; the third statement, which uses the **Run** method, does not require a reference to the Normal.dot project.




```
Normal.Module2.Macro1 
Call Normal.Module2.Macro1 
Application.Run MacroName:="Normal.Module2.Macro1"
```


## Example
<a name="sectionSection2"> </a>

This example prompts the user to enter a template name, module name, macro name, and parameter value, and then it runs that macro.


```
Dim strTemplate As String 
Dim strModule As String 
Dim strMacro As String 
Dim strParameter As String 
 
strTemplate = InputBox("Enter the template name") 
strModule = InputBox("Enter the module name") 
strMacro = InputBox("Enter the macro name") 
strParameter = InputBox("Enter a parameter value") 
Application.Run MacroName:=strTemplate &amp; "." _ 
 &amp; strModule &amp; "." &amp; strMacro, _ 
 varg1:=strParameter
```


## See also
<a name="sectionSection2"> </a>


#### Concepts


 [Application Object](d1cf6f8f-4e88-bf01-93b4-90a83f79cb44.md)
#### Other resources


 [Application Object Members](71669f1e-65f1-b0f1-b67d-355dfdbebe50.md)
