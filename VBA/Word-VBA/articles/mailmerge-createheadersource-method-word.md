---
title: MailMerge.CreateHeaderSource Method (Word)
keywords: vbawd10.chm153092198
f1_keywords:
- vbawd10.chm153092198
ms.prod: word
api_name:
- Word.MailMerge.CreateHeaderSource
ms.assetid: 607c668d-5f81-ecbe-d4c8-fbf509444683
ms.date: 06/08/2017
---


# MailMerge.CreateHeaderSource Method (Word)

Creates a Microsoft Word document that stores a header record that is used instead of the data source header record in a mail merge.


## Syntax

 _expression_ . **CreateHeaderSource**( **_Name_** , **_PasswordDocument_** , **_WritePasswordDocument_** , **_HeaderRecord_** )

 _expression_ Required. A variable that represents a **[MailMerge](mailmerge-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Name_|Required| **String**|The path and file name for the new header source.|
| _PasswordDocument_|Optional| **Variant**|The password required to open the new header source.|
| _WritePasswordDocument_|Optional| **Variant**|The password required to save changes to the new header source.|
| _HeaderRecord_|Optional| **Variant**|A string that specifies the field names for the header record. If this argument is omitted, the standard header record is used: "Title, FirstName, LastName, JobTitle, Company, Address1, Address2, City, State, PostalCode, Country, HomePhone, WorkPhone." To separate field names in Microsoft Windows, use the list separator specified in  **Regional Settings** in **Control Panel**.|

## Security

Avoid using hard-coded passwords in your applications. If a password is required in a procedure, request the password from the user, store it in a variable, and then use the variable in your code. For recommended best practices on how to do this, see [Security Notes for Microsoft Office Solution Developers](https://msdn.microsoft.com/en-us/library/office/ff860261.aspx). 


## Remarks

This method attaches the new header source to the specified document, which becomes a main document if it is not one already.The new header source uses a table to arrange mail merge field names. 


## Example

This example creates a header source with five field names and attaches the new header source named "Header.doc" to the active document.


```vb
ActiveDocument.MailMerge.CreateHeaderSource Name:="Header.doc", _ 
 HeaderRecord:="Name, Address, City, State, Zip"
```

This example creates a header source for the document named "Main.doc" (with the standard header record) and opens the data source named "Data.doc."




```vb
With Documents("Main.doc").MailMerge 
 .CreateHeaderSource Name:="Fields.doc" 
 .OpenDataSource Name:="C:\Documents\Data.doc" 
End With
```


## See also


#### Concepts


[MailMerge Object](mailmerge-object-word.md)

