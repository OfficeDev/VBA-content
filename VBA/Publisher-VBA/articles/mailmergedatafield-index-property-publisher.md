---
title: MailMergeDataField.Index Property (Publisher)
keywords: vbapb10.chm6422529
f1_keywords:
- vbapb10.chm6422529
ms.prod: publisher
api_name:
- Publisher.MailMergeDataField.Index
ms.assetid: f70d0266-0527-6871-632d-b45b617d75d4
ms.date: 06/08/2017
---


# MailMergeDataField.Index Property (Publisher)

Returns a  **Long** that represents the position of a particular item in a specified collection. .


## Syntax

 _expression_. **Index**

 _expression_A variable that represents a  **MailMergeDataField** object.


## Example

The following example loops through the  **MailMergeDataFields** collection and displays the **Index** and **Name** properties for each field.


```vb
Dim mmfLoop As MailMergeDataField 
 
With ActiveDocument.MailMerge.DataSource 
 If .DataFields.Count > 0 Then 
 For Each mmfLoop In .DataFields 
 Debug.Print "Field " &; mmfLoop.Name _ 
 &; " / Index " &; mmfLoop.Index 
 Next mmfLoop 
 Else 
 Debug.Print "No fields to report." 
 End If 
End With
```

The following example loops through the  **Plates** collection and displays the **Index** and **Name** properties for each plate.




```vb
Dim plaLoop As Plate 
 
If ActiveDocument.Plates.Count > 0 Then 
 For Each plaLoop In ActiveDocument.Plates 
 Debug.Print "Plate " &; plaLoop.Name _ 
 &; " / Index " &; plaLoop.Index 
 Next plaLoop 
Else 
 Debug.Print "No plates to report." 
End If
```


