---
title: Application.CloseCurrentDatabase Method (Access)
keywords: vbaac10.chm12506
f1_keywords:
- vbaac10.chm12506
ms.prod: access
api_name:
- Access.Application.CloseCurrentDatabase
ms.assetid: f5dec73c-54b4-c5ea-7cb9-25b5997f539e
ms.date: 06/08/2017
---


# Application.CloseCurrentDatabase Method (Access)

You can use the  **CloseCurrentDatabase** method to close the current database (either a Microsoft Access database or an Access project (.adp) from another application that has opened a database through Automation.


## Syntax

 _expression_. **CloseCurrentDatabase**

 _expression_ A variable that represents an **Application** object.


### Return Value

Nothing


## Remarks

For example, you might use this method from Microsoft Excel to close the database currently open in the Microsoft Access window before opening another database.

The  **CloseCurrentDatabase** method is useful when you have opened a Microsoft Access database from another application through Automation. Once you have created an instance of Microsoft Access from another application, you must also create a new database or specify an existing database to open. This database opens in the Microsoft Access window.

If you use the  **CloseCurrentDatabase** method to close the database that is open in the current instance of Microsoft Access, you can then open a different database without having to create another instance of Microsoft Access.


## Example

The following example opens a Microsoft Access database from another application through Automation, creates a new form and saves it, then closes the database.

You can enter this code in a Visual Basic module in any application that can act as a COM component. For example, you might run the following code from Microsoft Excel or Microsoft Visual Basic.

When the variable pointing to the  **Application** object goes out of scope, the instance of Microsoft Access that it represents closes as well. Therefore, you should declare this variable at the module level.




```vb
' Enter following in Declarations section of module. 
Dim appAccess As Access.Application 
 
Sub CreateForm() 
 Const strConPathToSamples = "C:\Program Files\Microsoft Office\Office12\Samples\" 
 
 Dim frm As Form, strDB As String 
 
 ' Initialize string to database path. 
 strDB = strConPathToSamples &; "Northwind.mdb" 
 ' Create new instance of Microsoft Access. 
 Set appAccess = CreateObject("Access.Application") 
 ' Open database in Microsoft Access window. 
 appAccess.OpenCurrentDatabase strDB 
 ' Create new form. 
 Set frm = appAccess.CreateForm 
 ' Save new form. 
 appAccess.DoCmd.Save , "NewForm1" 
 ' Close currently open database. 
 appAccess.CloseCurrentDatabase 
 Set AppAccess = Nothing 
End Sub
```


## See also


#### Concepts


[Application Object](application-object-access.md)

