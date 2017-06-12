---
title: Application.Run Method (Access)
keywords: vbaac10.chm12553
f1_keywords:
- vbaac10.chm12553
ms.prod: access
api_name:
- Access.Application.Run
ms.assetid: 4cdaf4cb-c25c-aaa4-96ab-52259f9f91c0
ms.date: 06/08/2017
---


# Application.Run Method (Access)

You can use the  **Run** method to carry out a specified Microsoft Access or user-defined **Function** or **Sub** procedure. **Variant**.


## Syntax

 _expression_. **Run**( ** _Procedure_**, ** _Arg1_**, ** _Arg2_**, ** _Arg3_**, ** _Arg4_**, ** _Arg5_**, ** _Arg6_**, ** _Arg7_**, ** _Arg8_**, ** _Arg9_**, ** _Arg10_**, ** _Arg11_**, ** _Arg12_**, ** _Arg13_**, ** _Arg14_**, ** _Arg15_**, ** _Arg16_**, ** _Arg17_**, ** _Arg18_**, ** _Arg19_**, ** _Arg20_**, ** _Arg21_**, ** _Arg22_**, ** _Arg23_**, ** _Arg24_**, ** _Arg25_**, ** _Arg26_**, ** _Arg27_**, ** _Arg28_**, ** _Arg29_**, ** _Arg30_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Procedure_|Required|**String**|The name of the  **Function** or **Sub** procedure to be run. If you are calling a procedure in another database use the project name and the procedure name separated by a dot in the form: " _projectname_. _procedurename_". If you execute Visual Basic code containing the  **Run** method in a library database, Microsoft Access looks for the procedure first in the library database, then in the current database.|
| _Arg1, Arg2, ...Arg30_|Optional|**Variant**|The arguments that should be passed to the  **Function** or **Sub** specified in the _Procedure_ argument.|

### Return Value

Variant


## Remarks

This method is useful when you are controlling Microsoft Access from another application through Automation, formerly called OLE Automation. For example, you can use the  **Run** method from an ActiveX component to carry out a **Sub** procedure that is defined within a Microsoft Access database.

You can set a reference to the Microsoft Access type library from any other ActiveX component and use the objects, methods, and properties defined in that library in your code. However, you can't set a reference to an individual Microsoft Access database from any application other than Microsoft Access.

For example, suppose you have defined a procedure named NewForm in a database with its  **ProjectName** property set to "WizCode." The NewForm procedure takes a string argument. You can call NewForm in the following manner from Visual Basic:




```vb
Dim appAccess As New Access.Application 
appAccess.OpenCurrentDatabase ("C:\My Documents\WizCode.mdb") 
appAccess.Run "WizCode.NewForm", "Some String"
```

If another procedure with the same name may reside in a different database, qualify the  _procedure_ argument, as shown in the preceding example, with the name of the database in which the desired procedure resides.

You can also use the  **Run** method to call a procedure in a referenced Microsoft Access database from another database.


## Example

The following example runs a user-defined  **Sub** procedure in a module in a Microsoft Access database from another application that acts as an Active X component.

To try this example, create a new database called WizCode.mdb and set its  **ProjectName** property to WizCode. Open a new module in that database and enter the following code. Save the module, and close the database.


 **Note**  You set the  **ProjectName** by selecting Tools, WizCode Properties... from the VBE main menu.




```vb
Public Sub Greeting(ByVal strName As String) 
 MsgBox ("Hello, " &; strName &; "!"), vbInformation, "Greetings" 
End Sub
```

Once you have completed this step, run the following code from Microsoft Excel or Microsoft Visual Basic. Make sure that you have added a reference to the Microsoft Access type library by clicking  **References** on the **Tools** menu and selecting **Microsoft Access 12.0 Object Library** in the **References** dialog box.




```vb
Private Sub RunAccessSub() 
 
 Dim appAccess As Access.Application 
 
 ' Create instance of Access Application object. 
 Set appAccess = CreateObject("Access.Application") 
 
 ' Open WizCode database in Microsoft Access window. 
 appAccess.OpenCurrentDatabase "C:\My Documents\WizCode.mdb", False 
 
 ' Run Sub procedure. 
 appAccess.Run "Greeting", "Joe" 
 Set appAccess = Nothing 
 
End Sub
```


## See also


#### Concepts


[Application Object](application-object-access.md)

