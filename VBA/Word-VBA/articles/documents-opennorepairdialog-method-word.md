---
title: Documents.OpenNoRepairDialog Method (Word)
keywords: vbawd10.chm158072852
f1_keywords:
- vbawd10.chm158072852
ms.prod: word
api_name:
- Word.Documents.OpenNoRepairDialog
ms.assetid: e299326e-dc8e-ab43-06fe-9b7625fb8beb
ms.date: 06/08/2017
---


# Documents.OpenNoRepairDialog Method (Word)

Opens the specified document and adds it to the [Documents](documents-object-word.md) collection.


## Syntax

 _expression_ . **OpenNoRepairDialog**( **_FileName_** , **_ConfirmConversions_** , **_ReadOnly_** , **_AddToRecentFiles_** , **_PasswordDocument_** , **_PasswordTemplate_** , **_Revert_** , **_WritePasswordDocument_** , **_WritePasswordTemplate_** , **_Format_** , **_Encoding_** , **_Visible_** , **_OpenAndRepair_** , **_DocumentDirection_** , **_NoEncodingDialog_** , **_XMLTransform_** )

 _expression_ A variable that represents a **[Documents](documents-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _FileName_|Required| **Variant**|The name of the document (paths are accepted).|
| _ConfirmConversions_|Optional| **Variant**| **True** to display the **Convert File** dialog box if the file is not in Microsoft Word format.|
| _ReadOnly_|Optional| **Variant**| **True** to open the document as read-only. This argument does not override the read-only recommended setting on a saved document. For example, if a document has been saved with read-only recommended turned on, setting the ReadOnly argument to **False** will not cause the file to be opened as read/write.|
| _AddToRecentFiles_|Optional| **Variant**| **True** to add the file name to the list of recently used files at the bottom of the **File** menu.|
| _PasswordDocument_|Optional| **Variant**|The password for opening the document.|
| _PasswordTemplate_|Optional| **Variant**|The password for opening the template.|
| _Revert_|Optional| **Variant**|Controls what happens if FileName is the name of an open document.  **True** to discard any unsaved changes to the open document and reopen the file. **False** to activate the open document.|
| _WritePasswordDocument_|Optional| **Variant**|The password for saving changes to the document.|
| _WritePasswordTemplate_|Optional| **Variant**|The password for saving changes to the template.|
| _Format_|Optional| **Variant**|The file converter to be used to open the document. Can be one of the  **[WdOpenFormat](wdopenformat-enumeration-word.md)** constants. The default is **wdOpenFormatAuto** .|
| _Encoding_|Optional| **Variant**|The document encoding (code page or character set) to be used by Word when you view the saved document. Can be any valid  **[MsoEncoding](http://msdn.microsoft.com/library/286bed6e-6028-a252-5e4f-b505234d9d34%28Office.15%29.aspx)** constant. For the list of valid **MsoEncoding** constants, see the Object Browser in the Visual Basic Editor. The default is the system code page.|
| _Visible_|Optional| **Variant**| **True** if the document is opened in a visible window. The default is **True** .|
| _OpenAndRepair_|Optional| **Variant**| **True** to repair the document to prevent document corruption.|
| _DocumentDirection_|Optional| **Variant**|Indicates the horizontal flow of text in a document. Can be any valid  **[WdDocumentDirection](wddocumentdirection-enumeration-word.md)** constant. The default is **wdLeftToRight** .|
| _NoEncodingDialog_|Optional| **Variant**| **True** to skip displaying the **Encoding** dialog box that Word displays if the text encoding cannot be recognized. The default is **False** .|
| _XMLTransform_|Optional| **Variant**|Specifies a transform to use.|

### Return Value

A  **[Document](document-object-word.md)** object that represents the specified document.


## Security

Avoid using hard-coded passwords in your applications. If a password is required in a procedure, request the password from the user, store it in a variable, and then use the variable in your code. For recommended best practices on how to do this, see [Security Notes for Microsoft Office Solution Developers](https://msdn.microsoft.com/en-us/library/office/ff860261.aspx). 


## Example

The following example opens MyDoc.doc as a read-only document.


```vb
Sub OpenDoc() 
 Documents.OpenNoRepairDialog FileName:="C:\MyFiles\MyDoc.doc", ReadOnly:=True 
End Sub
```

The following example opens Test.wp by using the WordPerfect 6.x file converter.




```vb
Sub OpenDoc2() 
 Dim fmt As Variant 
 fmt = Application.FileConverters("WordPerfect6x").OpenFormat 
 Documents.OpenNoRepairDialog FileName:="C:\MyFiles\Test.wp", Format:=fmt 
End Sub
```


## See also


#### Concepts


[Documents Collection Object](documents-object-word.md)

