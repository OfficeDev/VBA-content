---
title: MailMerge.OpenHeaderSource Method (Word)
keywords: vbawd10.chm153092209
f1_keywords:
- vbawd10.chm153092209
ms.prod: word
api_name:
- Word.MailMerge.OpenHeaderSource
ms.assetid: 0cf1102f-716b-4302-6d64-85fba29822ec
ms.date: 06/08/2017
---


# MailMerge.OpenHeaderSource Method (Word)

Attaches a mail merge header source to the specified document.


## Syntax

 _expression_ . **OpenHeaderSource**( **_Name_** , **_Format_** , **_ConfirmConversions_** , **_ReadOnly_** , **_AddToRecentFiles_** , **_PasswordDocument_** , **_PasswordTemplate_** , **_Revert_** , **_WritePasswordDocument_** , **_WritePasswordTemplate_** , **_OpenExclusive_** )

 _expression_ Required. A variable that represents a **[MailMerge](mailmerge-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Name_|Required| **String**|The file name of the header source.|
| _Format_|Optional| **Variant**|The file converter used to open the document. Can be one of the  **WdOpenFormat** constants. To specify an external file format, use the **OpenFormat** property with a **FileConverter** object to determine the value to use with this argument.|
| _ConfirmConversions_|Optional| **Variant**| **True** to display the **Convert File** dialog box if the file isn't in Microsoft Word format.|
| _ReadOnly_|Optional| **Variant**| **True** to open the header source on a read-only basis.|
| _AddToRecentFiles_|Optional| **Variant**| **True** to add the file name to the list of recently used files at the bottom of the **File** menu.|
| _PasswordDocument_|Optional| **Variant**|The password required to open the header source document. (See Remarks below.)|
| _PasswordTemplate_|Optional| **Variant**|The password required to open the header source template. (See Remarks below.)|
| _Revert_|Optional| **Variant**|Controls what happens if Name is the file name of an open document.  **True** to discard any unsaved changes to the open document and reopen the file; **False** to activate the open document.|
| _WritePasswordDocument_|Optional| **Variant**|The password required to save changes to the document data source. (See Remarks below.)|
| _WritePasswordTemplate_|Optional| **Variant**|The password required to save changes to the template data source. (See Remarks below.)|
| _OpenExclusive_|Optional| **Variant**| **True** to open exclusively.|

## Security

Avoid using hard-coded passwords in your applications. If a password is required in a procedure, request the password from the user, store it in a variable, and then use the variable in your code. For recommended best practices on how to do this, see [Security Notes for Microsoft Office Solution Developers](https://msdn.microsoft.com/en-us/library/office/ff860261.aspx). 


## Remarks

When a header source is attached, the first record in the header source is used in place of the header record in the data source.


## Example

This example sets the active document as a main document for form letters, and then it attaches the header source named "Header.doc" and the data document named "Names.doc."


```vb
With ActiveDocument.MailMerge 
 .MainDocumentType = wdFormLetters 
 .OpenHeaderSource Name:="C:\Documents\Header.doc", _ 
 Revert:=False, AddToRecentFiles:=False 
 .OpenDataSource Name:="C:\Documents\Names.doc" 
End With
```


## See also


#### Concepts


[MailMerge Object](mailmerge-object-word.md)

