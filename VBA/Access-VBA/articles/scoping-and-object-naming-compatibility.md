---
title: Scoping and Object-Naming Compatibility
keywords: vbaac10.chm5188026
f1_keywords:
- vbaac10.chm5188026
ms.prod: access
ms.assetid: 50e86279-78d0-c509-1598-250517153fe7
ms.date: 06/08/2017
---


# Scoping and Object-Naming Compatibility

Visual Basic scoping rules affect the names you choose for your objects, modules, and procedures.


## Modules and Other Objects with the Same Name

When you name a module, avoid prefacing module names with "Form_" or "Report_". Naming a module in this way could conflict with existing code you've written behind forms and reports.

If you have a module in an application created with version 1. _x_ or 2.0 of Access that doesn't follow these naming rules, Access generates an error when you try to convert your application. For example, a module named Form_Orders in an Access version 1. _x_ or 2.0 database would generate an error and you would be asked to rename the module before attempting to convert it.


## Modules and Procedures with the Same Name

Although it is not suggested, you can have a procedure with the same name as a module. To call that procedure from an expression anywhere in your application, you must use a fully qualified name for the procedure, including both the module name and the procedure name, as in the following example:


```vb
IsLoaded.IsLoaded("Orders")
```


 **Note**  This will not work with the  **Runcode** action in macros. Accessing procedures with the same name as a module is not possible with macros.


## Procedures and Controls with the Same Name

If you call a procedure from a form, and that procedure has the same name as a control on the form, you must fully qualify the procedure call with the name of the module in which it resides. For example, if you want to call a procedure named PrintInvoice that resides in a standard module named Utilities, and there's also a button on the same form named PrintInvoice, use the fully qualified name when you call the procedure from your form or form module.


## Controls with Similar Names

You can't have a control with a name that differs from an existing control's name by only a space or a symbol. For example, if you have a control named [Last_Name], you can't have a control named [Last Name] or [Last+Name].


## Modules with the Same Names as Type Libraries

You can't save a module with the same name as a type library. If you try to save a module with the name "ADO", "Access", "DAO" or "VBA", you'll get an error stating that the name conflicts with an existing module, project, or object library. Similarly, if you've set a reference to another type library, such as the Excel type library, you can't save a module with the name "Excel".


## Fields with the Same Names as Methods

If a field in the table has the same name as an ActiveX Data Objects (ADO) method on an ADO  **Recordset** object, or a Data Access Object (DAO) method on a DAO **Recordset** object, you can't refer to the corresponding field in the recordset with the . (dot) syntax. You must use the ! (exclamation point) syntax, or Access will generate an error. The following example shows how to refer to a field called AddNew in a recordset opened on a table called Contacts:


## 


```vb
Dim rst As New ADODB.Recordset 
rst.Open "Contacts",CurrentProject.Connection, _ 
 adOpenKeySet,adLockOptimistic 
Debug.Print rst!AddNew 

```


## 


```vb
Dim dbs As Database, rst As DAO.Recordset 
Set dbs = CurrentDb 
Set rst = dbs.OpenRecordset("Contacts") 
Debug.Print rst!AddNew
```


## Modules with the Same Names as Visual Basic Functions

If you save a module with the same name as an intrinsic Visual Basic function, Access will generate an error when you try to run that function. For example, if you save a module named MsgBox, and then try to run a procedure that calls the  **MsgBox** function, Access generates the error "Expected variable or procedure, not module."


## Modules with the Same Names as Objects

If a database created with a previous version of Access includes a module that has the same name as an Access object, an ADO object, or a DAO object, you may encounter compilation errors when you convert your database. For example, a module named "Form" or "Database" may generate a compilation error. To avoid these errors, rename the module.


## Naming Fields Used in Expressions or Bound to Controls on Forms and Reports

When you create a field in a table that will be bound to a control on a report or used in an expression in the  **ControlSource** property of a control or a report, avoid assigning the field a name that's the same name as a method of the **[Application](application-object-access.md)** object. To see a list of methods of the **Application** object, click **Object Browser** on the **View** menu while in module Design view. Click **Access** in the **Project/Library** box, click **Application** in the **Classes** box, and view the methods of the **Application** object in the **Members Of** box.

When you create a field in a table that will be bound to a control on a form or report, avoid assigning the field any of the following names: AddRef, GetIDsOfNames, GetTypeInfo, GetTypeInfoCount, Invoke, QueryInterface, or Release.


## Identifiers with Same Names as Visual Basic Keywords

The version of Visual Basic that's used by Access 97 (and later) contains some new Visual Basic keywords, so you can no longer use these keywords as identifiers. These keywords are:  **AddressOf**, **Assert**, **Decimal**, **DefDec**, **Enum**, **Event**, **Friend**, **Implements**, **RaiseEvent**, and **WithEvents**. When you convert a database developed with a prior version of Access, existing identifiers that are the same as a new Visual Basic keyword will cause a compile error. To correct this problem, rename the identifiers.


## Project Names the Same as Access Objects

A project name is the string that is the name of your Access application. In prior versions of Access, the project name was the name of the database. Beginning in Access 2000, the project name is specified by the  **ProjectName** property setting and its default setting is the name of the database. If you convert a database with a name that is the same as a class of objects, for example, "application," "form," or "report," Access will append an underscore character to the database name to create a project name that does not conflict with existing objects.


