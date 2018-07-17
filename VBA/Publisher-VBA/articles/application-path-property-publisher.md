---
title: Application.Path Property (Publisher)
keywords: vbapb10.chm131097
f1_keywords:
- vbapb10.chm131097
ms.prod: publisher
api_name:
- Publisher.Application.Path
ms.assetid: 36ac9a9c-8235-aeba-c3d5-d39aef960cc5
ms.date: 06/08/2017
---


# Application.Path Property (Publisher)

Returns a  **String** indicating the full path to the file of the saved active publication, not including the last separator or file name.


## Syntax

 _expression_. **Path**

 _expression_A variable that represents an  **Application** object.


## Remarks

The  **[FullName](document-fullname-property-publisher.md)** property can be used to return both the path and file name.


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


## See also


#### Concepts


 [Application Object](application-object-publisher.md)

