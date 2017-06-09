---
title: Application.OpenCurrentDatabase Method (Access)
keywords: vbaac10.chm12601
f1_keywords:
- vbaac10.chm12601
ms.prod: access
api_name:
- Access.Application.OpenCurrentDatabase
ms.assetid: fd214849-02ac-eaa6-7525-9aee42b92f3d
ms.date: 06/08/2017
---


# Application.OpenCurrentDatabase Method (Access)

You can use the  **OpenCurrentDatabase** method to open an existing Microsoft Access database as the current database.


## Syntax

 _expression_. **OpenCurrentDatabase**( ** _filepath_**, ** _Exclusive_**, ** _bstrPassword_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _filepath_|Required|**String**|The name of an existing database file, including the path name and the file name extension.|
| _Exclusive_|Optional|**Boolean**|Specifies whether you want to open the database in exclusive mode. The default value is  **False**, which specifies that the database should be opened in shared mode.|
| _bstrPassword_|Optional|**String**|The password to open the specified database.|

### Return Value

Nothing


## Remarks

You can use this method to open a database from another application that is controlling Microsoft Access through Automation, formerly called OLE Automation. For example, you can use the  **OpenCurrentDatabase** method from Microsoft Excel to open the Northwind.mdb sample database in the Microsoft Access window. Once you have created an instance of Microsoft Access from another application, you must also create a new database or specify a particular database to open. This database opens in the Microsoft Access window.

If you have already opened a database and wish to open another database in the Microsoft Access window, you can use the  **[CloseCurrentDatabase](application-closecurrentdatabase-method-access.md)** method to close the first database before opening another.




 **Note**  Use the  **[OpenAccessProject](application-openaccessproject-method-access.md)** method to open an existing Microsoft Access project (.adp) as the current database.




 **Note**  Don't confuse the  **OpenCurrentDatabase** method with the ActiveX Data Objects (ADO) **Open** method or the Data Access Object (DAO) **OpenDatabase** method. The **OpenCurrentDatabase** method opens a database in the Microsoft Access window. The DAO **OpenDatabase** method returns a **Database** object variable, which represents a particular database but don't actually open that database in the Microsoft Access window.


## Example

The following example opens a Microsoft Access database from another application through Automation and then opens a form in that database.

You can enter this code in a Visual Basic module in any application that can act as a COM component. For example, you might run the following code from Microsoft Excel, Microsoft Visual Basic, or Microsoft Access.

When the variable pointing to the  **Application** object goes out of scope, the instance of Microsoft Access that it represents closes as well. Therefore, you should declare this variable at the module level.




```vb
' Include the following in Declarations section of module. 
Dim appAccess As Access.Application 
 
Sub DisplayForm() 
 
 Dim strDB as String 
 
 ' Initialize string to database path. 
 Const strConPathToSamples = "C:\Program " _ 
 &; "Files\Microsoft Office\Office11\Samples\" 
 
 strDB = strConPathToSamples &; "Northwind.mdb" 
 ' Create new instance of Microsoft Access. 
 Set appAccess = _ 
 CreateObject("Access.Application") 
 ' Open database in Microsoft Access window. 
 appAccess.OpenCurrentDatabase strDB 
 ' Open Orders form. 
 appAccess.DoCmd.OpenForm "Orders" 
End Sub
```


## See also


#### Concepts


[Application Object](application-object-access.md)

