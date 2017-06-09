---
title: Subdocuments.AddFromFile Method (Word)
keywords: vbawd10.chm159907940
f1_keywords:
- vbawd10.chm159907940
ms.prod: word
api_name:
- Word.Subdocuments.AddFromFile
ms.assetid: 7f9e73a9-bea9-815e-eccc-3406e6d5dd63
ms.date: 06/08/2017
---


# Subdocuments.AddFromFile Method (Word)

Adds the specified subdocument to the master document at the start of the selection and returns a  **Subdocument** object.


## Syntax

 _expression_ . **AddFromFile**( **_Name_** , **_ConfirmConversions_** , **_ReadOnly_** , **_PasswordDocument_** , **_PasswordTemplate_** , **_Revert_** , **_WritePasswordDocument_** , **_WritePasswordTemplate_** )

 _expression_ Required. A variable that represents a **[Subdocuments](subdocuments-object-word.md)** collection.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Name_|Required| **String**|The file name of the subdocument to be inserted into the master document.|
| _ConfirmConversions_|Optional| **Variant**| **True** to confirm file conversion in the **Convert File** dialog box if the file isn't in Word format.|
| _ReadOnly_|Optional| **Variant**| **True** to insert the subdocument as a read-only document.|
| _PasswordDocument_|Optional| **Variant**|The password required to open the subdocument if it is password protected.|
| _PasswordTemplate_|Optional| **Variant**|The password required to open the template attached to the subdocument if the template is password protected.|
| _Revert_|Optional| **Variant**|Controls what happens if Name is the file name of an open document.  **True** to insert the saved version of the subdocument. **False** to insert the open version of the subdocument, which may contain unsaved changes.|
| _WritePasswordDocument_|Optional| **Variant**|The password required to save changes to the document file if it is write-protected.|
| _WritePasswordTemplate_|Optional| **Variant**|The password required to save changes to the template attached to the subdocument if the template is write-protected.|

### Return Value

Subdocument


## Remarks

If the active view isn't either outline view or master document view, an error occurs.


## Example

This example adds a subdocument named "Subdoc.doc" to the active document.


```vb
ActiveDocument.ActiveWindow.View.Type = wdMasterView 
ActiveDocument.Subdocuments.AddFromFile _ 
 Name:="C:\Subdoc.doc"
```

This example adds a password-protected subdocument named "Subdoc.doc" to the active document on a read-only basis and sets the PasswordDocument parameter to a String variable.




```
Selection.Range.Subdocuments.AddFromFile Name:="C:\Subdoc.doc", _ 
 ReadOnly:=True, PasswordDocument:=strPassword
```


## See also


#### Concepts


[Subdocuments Collection Object](subdocuments-object-word.md)

