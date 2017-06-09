---
title: MailMerge.CreateDataSource Method (Word)
keywords: vbawd10.chm153092197
f1_keywords:
- vbawd10.chm153092197
ms.prod: word
api_name:
- Word.MailMerge.CreateDataSource
ms.assetid: 720beea6-3496-c760-3465-117ee4beffb1
ms.date: 06/08/2017
---


# MailMerge.CreateDataSource Method (Word)

Creates a Microsoft Word document that uses a table to store data for a mail merge.


## Syntax

 _expression_ . **CreateDataSource**( **_Name_** , **_PasswordDocument_** , **_WritePasswordDocument_** , **_HeaderRecord_** , **_MSQuery_** , **_SQLStatement_** , **_SQLStatement1_** , **_Connection_** , **_LinkToSource_** )

 _expression_ Required. A variable that represents a **[MailMerge](mailmerge-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Name_|Optional| **Variant**|The path and file name for the new data source.|
| _PasswordDocument_|Optional| **Variant**|The password required to open the new data source.|
| _WritePasswordDocument_|Optional| **Variant**|The password required to save changes to the data source.|
| _HeaderRecord_|Optional| **Variant**|Field names for the header record. If this argument is omitted, the standard header record is used: "Title, FirstName, LastName, JobTitle, Company, Address1, Address2, City, State, PostalCode, Country, HomePhone, WorkPhone." To separate field names, use the list separator specified in  **Regional Settings** in **Control Panel**.|
| _MSQuery_|Optional| **Variant**| **True** to launch Microsoft Query, if it is installed. The Name, PasswordDocument, and HeaderRecord arguments are ignored.|
| _SQLStatement_|Optional| **Variant**|Defines query options for retrieving data.|
| _SQLStatement1_|Optional| **Variant**|If the query string is longer than 255 characters, SQLStatement specifies the first portion of the string, and SQLStatement1 specifies the second portion.|
| _Connection_|Optional| **Variant**|A range within which the query specified by SQLStatement will be performed.|
| _LinkToSource_|Optional| **Variant**| **True** to perform the query specified by Connection and SQLStatement each time the main document is opened.|

## Security

Avoid using hard-coded passwords in your applications. If a password is required in a procedure, request the password from the user, store it in a variable, and then use the variable in your code. For recommended best practices on how to do this, see [Security Notes for Microsoft Office Solution Developers](https://msdn.microsoft.com/en-us/library/office/ff860261.aspx). 


## Remarks

When you use the  **CreateDataSource** method, Word attaches the new data source to the specified document, which becomes a main document, if it is not one already.

How you specify the range for the Connection argument depends on how data is retrieved. For example:


- When retrieving data through ODBC, you specify a connection string.
    
- When retrieving data from Microsoft Office Excel using dynamic data exchange (DDE), you specify a named range. 
 **Security Note**  


    
- When retrieving data from Microsoft Office Access, you specify the word "Table" or "Query" followed by the name of a table or query.
    

## Example

This example creates a new data source document named "Data.doc" and attaches the data source to the active document. The new data source includes a five-column table that has the field names specified by the HeaderRecord argument.


```vb
ActiveDocument.MailMerge.CreateDataSource _ 
 Name:="C:\Documents\Data.doc", _ 
 HeaderRecord:="Name, Address, City, State, Zip"
```


## See also


#### Concepts


[MailMerge Object](mailmerge-object-word.md)

