---
title: FileConverter.SaveFormat Property (Word)
keywords: vbawd10.chm161021954
f1_keywords:
- vbawd10.chm161021954
ms.prod: word
api_name:
- Word.FileConverter.SaveFormat
ms.assetid: d837cd22-38eb-5160-1f85-16001448213e
ms.date: 06/08/2017
---


# FileConverter.SaveFormat Property (Word)

Returns the file format of the specified document or file converter. Read-only  **Long** .


## Syntax

 _expression_ . **SaveFormat**

 _expression_ Required. A variable that represents a **[FileConverter](fileconverter-object-word.md)** object.


## Remarks

This property returns a unique number that specifies an external file converter or a  **[WdSaveFormat](wdsaveformat-enumeration-word.md)** constant. Use the value of the **SaveFormat** property for the _FileFormat_ argument of the **[SaveAs2](document-saveas2-method-word.md)** method to save a document in a file format for which there isn't a corresponding **WdSaveFormat** constant.


## Example

This example creates a new document and lists in a table the converters that can be used to save documents and their corresponding  **SaveFormat** values.


```vb
Sub FileConverterList() 
 Dim cnvFile As FileConverter 
 Dim docNew As Document 
 
 'Create a new document and set a tab stop 
 Set docNew = Documents.Add 
 docNew.Paragraphs.Format.TabStops.Add _ 
 Position:=InchesToPoints(3) 
 
 'List all the converters in the FileConverters collection 
 With docNew.Content 
 .InsertAfter "Name" &; vbTab &; "Number" 
 .InsertParagraphAfter 
 For Each cnvFile In FileConverters 
 If cnvFile.CanSave = True Then 
 .InsertAfter cnvFile.FormatName &; vbTab &; _ 
 cnvFile.SaveFormat 
 .InsertParagraphAfter 
 End If 
 Next 
 .ConvertToTable 
 End With 
 
End Sub
```

This example saves the active document in the WordPerfect 5.1 or 5.2 secondary file format.




```vb
ActiveDocument.SaveAs _ 
 FileFormat:=FileConverters("WrdPrfctDat").SaveFormat
```


## See also


#### Concepts


[FileConverter Object](fileconverter-object-word.md)

