---
title: MailMerge.OpenDataSource Method (Word)
keywords: vbawd10.chm153092208
f1_keywords:
- vbawd10.chm153092208
ms.prod: word
api_name:
- Word.MailMerge.OpenDataSource
ms.assetid: fea43151-bb56-34ad-090c-24d9e47aeaac
ms.date: 06/08/2017
---


# MailMerge.OpenDataSource Method (Word)

Attaches a data source to the specified document, which becomes a main document if it is not one already.


## Syntax

 _expression_ . **OpenDataSource**( **_Name_** , **_Format_** , **_ConfirmConversions_** , **_ReadOnly_** , **_LinkToSource_** , **_AddToRecentFiles_** , **_PasswordDocument_** , **_PasswordTemplate_** , **_Revert_** , **_WritePasswordDocument_** , **_WritePasswordTemplate_** , **_Connection_** , **_SQLStatement_** , **_SQLStatement1_** , **_OpenExclusive_** , **_SubType_** )

 _expression_ Required. A variable that represents a **[MailMerge](mailmerge-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Name_|Required| **String**|The data source file name. You can specify a Microsoft Query (.qry) file instead of specifying a data source, a connection string, and a query string.|
| _Format_|Optional| **Variant**|The file converter used to open the document. Can be one of the  **WdOpenFormat** constants. To specify an external file format, use the **OpenFormat** property with the **FileConverter** object to determine the value to use with this argument.|
| _ConfirmConversions_|Optional| **Variant**| **True** to display the **Convert File** dialog box if the file is not in Microsoft Word format.|
| _ReadOnly_|Optional| **Variant**| **True** to open the data source on a read-only basis.|
| _LinkToSource_|Optional| **Variant**| **True** to perform the query specified by Connection and SQLStatement each time the main document is opened.|
| _AddToRecentFiles_|Optional| **Variant**| **True** to add the file name to the list of recently used files at the bottom of the **File** menu.|
| _PasswordDocument_|Optional| **Variant**|The password used to open the data source. (See Remarks below.)|
| _PasswordTemplate_|Optional| **Variant**|The password used to open the template. (See Remarks below.)|
| _Revert_|Optional| **Variant**|Controls what happens if Name is the file name of an open document.  **True** to discard any unsaved changes to the open document and reopen the file; **False** to activate the open document.|
| _WritePasswordDocument_|Optional| **Variant**|The password used to save changes to the document. (See Remarks below.)|
| _WritePasswordTemplate_|Optional| **Variant**|The password used to save changes to the template. (See Remarks below.)|
| _Connection_|Optional| **Variant**|A range within which the query specified by SQLStatement is to be performed. (See Remarks below.) |
| _SQLStatement_|Optional| **Variant**|Defines query options for retrieving data. (See Remarks below.)|
| _SQLStatement1_|Optional| **Variant**|If the query string is longer than 255 characters, SQLStatement specifies the first portion of the string, and SQLStatement1 specifies the second portion. (See Remarks below.)|
| _OpenExclusive_|Optional| **Variant**| **True** to open exclusively.|
| _SubType_|Optional| **Variant**|Can be one of the  **WdMergeSubType** constants.|

## Remarks

To determine the ODBC connection and query strings, set query options manually and use the  **QueryString** property to return the connection string. The following table includes some commonly used SQL keywords.



|**Keyword**|**Description**|
|:-----|:-----|
|DSN|The name of the ODBC data source|
|UID|The user logon ID|
|PWD|The user-specified password|
|DBQ|The database file name|
|FIL|The file type|
To instruct Word to use the same connection method as in earlier versions of Word (Dynamic Data Exchange (DDE) for Microsoft Office Access and Microsoft Office Excel data sources) use  `SubType:=wdMergeSubTypeWord2000`.

How you specify the range depends on how data is retrieved. For example:


- When retrieving data through Open Database Connectivity (ODBC), you specify a connection string.
    
- When retrieving data from Excel using dynamic data exchange (DDE), you specify a named range.
 **Security Note**  


    
- When retrieving data from Access, you specify the word "Table" or "Query" followed by the name of a table or query.
    

 **Security Note**  




 **Security Note**  



Avoid using hard-coded passwords in your applications. If a password is required in a procedure, request the password from the user, store it in a variable, and then use the variable in your code. For recommended best practices on how to do this, see [Security Notes for Microsoft Office Solution Developers](https://msdn.microsoft.com/en-us/library/office/ff860261.aspx). 


## Example

This example creates a new main document and attaches the Orders table from an Access database named "Northwind.mdb."


```vb
Dim docNew As Document 
 
Set docNew = Documents.Add 
 
With docNew.MailMerge 
 .MainDocumentType = wdFormLetters 
 .OpenDataSource _ 
 Name:="C:\Program Files\Microsoft Office" &; _ 
 "\Office\Samples\Northwind.mdb", _ 
 LinkToSource:=True, AddToRecentFiles:=False, _ 
 Connection:="TABLE Orders" 
End With
```

This example creates a new main document and attaches the Excel worksheet named ?Names.xls.? The Connection argument retrieves data from the range named "Sales."




```vb
Dim docNew As Document 
 
Set docNew = Documents.Add 
 
With docNew.MailMerge 
 .MainDocumentType = wdCatalog 
 .OpenDataSource Name:="C:\Documents\Names.xls", _ 
 ReadOnly:=True, _ 
 Connection:="Sales" 
End With
```

This example uses ODBC to attach the Access database named "Northwind.mdb" to the active document. The SQLStatement argument selects the records in the Customers table.




```vb
Dim strConnection As String 
 
With ActiveDocument.MailMerge 
 .MainDocumentType = wdFormLetters 
 strConnection = "DSN=MS Access Databases;" _ 
 &; "DBQ=C:\Northwind.mdb;" _ 
 &; "FIL=RedISAM;" 
 .OpenDataSource Name:="C:\NorthWind.mdb", _ 
 Connection:=strConnection, _ 
 SQLStatement:="SELECT * FROM Customers" 
End With
```


## See also


#### Concepts


[MailMerge Object](mailmerge-object-word.md)

