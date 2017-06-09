---
title: Document.SaveAs2 Method (Word)
keywords: vbawd10.chm158007864
f1_keywords:
- vbawd10.chm158007864
ms.prod: word
api_name:
- Word.SaveAs2
ms.assetid: aa491007-0e31-26f5-3a5e-477381529b6e
ms.date: 06/08/2017
---


# Document.SaveAs2 Method (Word)

Saves the specified document with a new name or format. Some of the arguments for this method correspond to the options in the  **Save As** dialog box ( **File** tab).


## Syntax

 _expression_ . **SaveAs2**( **_FileName_** , **_FileFormat_** , **_LockComments_** , **_Password_** , **_AddToRecentFiles_** , **_WritePassword_** , **_ReadOnlyRecommended_** , **_EmbedTrueTypeFonts_** , **_SaveNativePictureFormat_** , **_SaveFormsData_** , **_SaveAsAOCELetter_** , **_Encoding_** , **_InsertLineBreaks_** , **_AllowSubstitutions_** , **_LineEnding_** , **_AddBiDiMarks_** , **_CompatibilityMode_** )

 _expression_ An expression that returns a **[Document](document-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _FileName_|Optional| **Variant**|The name for the document. The default is the current folder and file name. If the document has never been saved, the default name is used (for example, Doc1.doc). If a document with the specified file name already exists, the document is overwritten without prompting the user.|
| _FileFormat_|Optional| **Variant**|The format in which the document is saved. Can be any  **[WdSaveFormat](wdsaveformat-enumeration-word.md)** constant. To save a document in another format, specify the appropriate value for the **[SaveFormat](fileconverter-saveformat-property-word.md)** property of the **[FileConverter](fileconverter-object-word.md)** object.|
| _LockComments_|Optional| **Variant**| **True** to lock the document for comments. The default is **False** .|
| _Password_|Optional| **Variant**|A password string for opening the document. (See Remarks below.)|
| _AddToRecentFiles_|Optional| **Variant**| **True** to add the document to the list of recently used files on the **File** menu. The default is **True** .|
| _WritePassword_|Optional| **Variant**|A password string for saving changes to the document. (See Remarks below.)|
| _ReadOnlyRecommended_|Optional| **Variant**| **True** to have Microsoft Word suggest read-only status whenever the document is opened. The default is **False** .|
| _EmbedTrueTypeFonts_|Optional| **Variant**| **True** to save TrueType fonts with the document. If omitted, the EmbedTrueTypeFonts argument assumes the value of the **[EmbedTrueTypeFonts](document-embedtruetypefonts-property-word.md)** property.|
| _SaveNativePictureFormat_|Optional| **Variant**|If graphics were imported from another platform (for example, Macintosh),  **True** to save only the Microsoft Windows version of the imported graphics.|
| _SaveFormsData_|Optional| **Variant**| **True** to save the data entered by a user in a form as a record.|
| _SaveAsAOCELetter_|Optional| **Variant**|If the document has an attached mailer,  **True** to save the document as an AOCE letter (the mailer is saved).|
| _Encoding_|Optional| **Variant**|The code page, or character set, to use for documents saved as encoded text files. The default is the system code page. You cannot use all  **[MsoEncoding](http://msdn.microsoft.com/library/286bed6e-6028-a252-5e4f-b505234d9d34%28Office.15%29.aspx)** constants with this parameter.|
| _InsertLineBreaks_|Optional| **Variant**|If the document is saved as a text file,  **True** to insert line breaks at the end of each line of text.|
| _AllowSubstitutions_|Optional| **Variant**|If the document is saved as a text file,  **True** allows Word to replace some symbols with text that looks similar. For example, displaying the copyright symbol as (c). The default is **False** .|
| _LineEnding_|Optional| **Variant**|The way Word marks the line and paragraph breaks in documents saved as text files. Can be one of the following  **[WdLineEndingType](wdlineendingtype-enumeration-word.md)** constants: **wdCRLF** (default) or **wdCROnly** .|
| _AddBiDiMarks_|Optional| **Variant**| **True** adds control characters to the output file to preserve bi-directional layout of the text in the original document.|
| _CompatibilityMode_|Optional| **Variant**|The compatibility mode that Word uses when opening the document.  **[WdCompatibilityMode](wdcompatibilitymode-enumeration-word.md)** constant.<table><tr><th>**Important**</th></tr><tr><td>By default, if no value is specified for this parameter, Word enters a value of 0, which specifies that the current compatibility mode of the document should be retained.</td></tr></table>|

### Return Value

Nothing


## Example

The following code example saves the active document as Test.rtf in rich-text format (RTF).


```vb
Sub SaveAsRTF() 
    ActiveDocument.SaveAs2 FileName:="Text.rtf", _ 
        FileFormat:=wdFormatRTF 
End Sub
```

The following code example saves the active document in text-file format with the extension ".txt".




```vb
Sub SaveAsTextFile() 
    Dim strDocName As String 
    Dim intPos As Integer 
 
    ' Find position of extension in file name 
    strDocName = ActiveDocument.Name 
    intPos = InStrRev(strDocName, ".") 
 
    If intPos = 0 Then 
 
        ' If the document has not yet been saved 
        ' Ask the user to provide a file name 
        strDocName = InputBox("Please enter the name " &; _ 
            "of your document.") 
    Else 
 
        ' Strip off extension and add ".txt" extension 
        strDocName = Left(strDocName, intPos - 1) 
        strDocName = strDocName &; ".txt" 
    End If 
 
    ' Save file with new extension 
    ActiveDocument.SaveAs2 FileName:=strDocName, _ 
        FileFormat:=wdFormatText 
End Sub
```

The following code example loops through all the installed converters and, if it finds the WordPerfect 6.0 converter, it saves the active document using the converter.




```vb
Sub SaveWithConverter() 
 
    Dim cnvWrdPrf As FileConverter 
 
    ' Look for WordPerfect file converter 
    ' And save document using the converter 
    ' For the FileFormat converter value 
    For Each cnvWrdPrf In Application.FileConverters 
        If cnvWrdPrf.ClassName = "WrdPrfctWin" Then 
            ActiveDocument.SaveAs2 FileName:="MyWP.doc", _ 
                FileFormat:=cnvWrdPrf.SaveFormat 
        End If 
    Next cnvWrdPrf 
 
End Sub
```

The following code example shows a procedure that saves a document with a password.




```vb
Sub SaveWithPassword(docCurrent As Document, strPWD As String) 
    With docCurrent 
        .SaveAs2 WritePassword:=strPWD 
    End With 
End Sub
```


## See also


#### Concepts


[Document Object](document-object-word.md)

