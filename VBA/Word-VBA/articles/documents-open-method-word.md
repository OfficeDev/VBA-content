---
title: Documents.Open Method (Word)
keywords: vbawd10.chm158072851
f1_keywords:
- vbawd10.chm158072851
ms.prod: word
api_name:
- Word.Documents.Open
ms.assetid: 9e61e9d5-58d1-833a-5f93-b87299deb400
ms.date: 06/08/2017
---


# Documents.Open Method (Word)

Opens the specified document and adds it to the  **Documents** collection. Returns a **Document** object.


## Syntax

 _expression_ . **Open**( **_FileName_** , **_ConfirmConversions_** , **_ReadOnly_** , **_AddToRecentFiles_** , **_PasswordDocument_** , **_PasswordTemplate_** , **_Revert_** , **_WritePasswordDocument_** , **_WritePasswordTemplate_** , **_Format_** , **_Encoding_** , **_Visible_** , **_OpenConflictDocument_** , **_OpenAndRepair_** , **_DocumentDirection_** , **_NoEncodingDialog_** )

 _expression_ Required. A variable that represents a **[Documents](documents-object-word.md)** collection.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _FileName_|Required| **Variant**|The name of the document (paths are accepted).|
| _ConfirmConversions_|Optional| **Variant**| **True** to display the **Convert File** dialog box if the file isn't in Microsoft Word format.|
| _ReadOnly_|Optional| **Variant**| **True** to open the document as read-only. This argument doesn't override the read-only recommended setting on a saved document. For example, if a document has been saved with read-only recommended turned on, setting the ReadOnly argument to **False** will not cause the file to be opened as read/write.|
| _AddToRecentFiles_|Optional| **Variant**| **True** to add the file name to the list of recently used files at the bottom of the **File** menu.|
| _PasswordDocument_|Optional| **Variant**|The password for opening the document.|
| _PasswordTemplate_|Optional| **Variant**|The password for opening the template.|
| _Revert_|Optional| **Variant**|Controls what happens if FileName is the name of an open document.  **True** to discard any unsaved changes to the open document and reopen the file. **False** to activate the open document.|
| _WritePasswordDocument_|Optional| **Variant**|The password for saving changes to the document.|
| _WritePasswordTemplate_|Optional| **Variant**|The password for saving changes to the template.|
| _Format_|Optional| **Variant**|The file converter to be used to open the document. Can be one of the  **WdOpenFormat** constants. The default value is **wdOpenFormatAuto** . To specify an external file format, apply the **OpenFormat** property to a **FileConverter** object to determine the value to use with this argument.|
| _Encoding_|Optional| **Variant**|The document encoding (code page or character set) to be used by Microsoft Word when you view the saved document. Can be any valid  **MsoEncoding** constant. For the list of valid **MsoEncoding** constants, see the Object Browser in the Visual Basic Editor. The default value is the system code page.|
| _Visible_|Optional| **Variant**| **True** if the document is opened in a visible window. The default value is **True** .|
| _OpenConflictDocument_|Optional| **Variant**|Specifies whether to open the conflict file for a document with an offline conflict.|
| _OpenAndRepair_|Optional| **Variant**| **True** to repair the document to prevent document corruption.|
| _DocumentDirection_|Optional| **WdDocumentDirection**|Indicates the horizontal flow of text in a document. The default value is  **wdLeftToRight** .|
| _NoEncodingDialog_|Optional| **Variant**| **True** to skip displaying the Encoding dialog box that Word displays if the text encoding cannot be recognized. The default value is **False** .|

### Return Value

Document


## Security

Avoid using hard-coded passwords in your applications. If a password is required in a procedure, request the password from the user, store it in a variable, and then use the variable in your code. For recommended best practices on how to do this, see [Security Notes for Microsoft Office Solution Developers](https://msdn.microsoft.com/en-us/library/office/ff860261.aspx). 


## Example

This example opens MyDoc.doc as a read-only document.


```vb
Sub OpenDoc() 
 Documents.Open FileName:="C:\MyFiles\MyDoc.doc", ReadOnly:=True 
End Sub
```

This example opens Test.wp using the WordPerfect 6.x file converter.




```vb
Sub OpenDoc2() 
 Dim fmt As Variant 
 fmt = Application.FileConverters("WordPerfect6x").OpenFormat 
 Documents.Open FileName:="C:\MyFiles\Test.wp", Format:=fmt 
End Sub
```


## See also


#### Concepts


[Documents Collection Object](documents-object-word.md)

