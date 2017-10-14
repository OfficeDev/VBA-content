---
title: Document.FullName Property (Publisher)
keywords: vbapb10.chm196625
f1_keywords:
- vbapb10.chm196625
ms.prod: publisher
api_name:
- Publisher.Document.FullName
ms.assetid: 137e4310-8431-ed2a-503a-c225378a9a74
ms.date: 06/08/2017
---


# Document.FullName Property (Publisher)

Returns a  **String** representing the full file name of the saved active publication, including its path and file name. Read-only.


## Syntax

 _expression_. **FullName**

 _expression_A variable that represents a  **Document** object.


### Return Value

String


## Remarks

The  **FullName** property can be used to return both path and file name as returned by the **[Path](document-path-property-publisher.md)** and **[Name](document-name-property-publisher.md)** properties.


## Example

The following example demonstrates the differences between the  **Path**,  **Name**, and  **FullName** properties. This example is best illustrated if the publication is saved in a folder other than the default.


```vb
Sub PathNames() 
 
 Dim strPath As String 
 Dim strName As String 
 Dim strFullName As String 
 
 strPath = Application.ActiveDocument.Path 
 strName = Application.ActiveDocument.Name 
 strFullName = Application.ActiveDocument.FullName 
 
 ' Note the file name &; path differences 
 ' while executing. 
 MsgBox "The path is: " &; strPath 
 MsgBox "The file name is: " &; strName 
 MsgBox "The path &; file name are: " &; strFullName 
 
End Sub
```


