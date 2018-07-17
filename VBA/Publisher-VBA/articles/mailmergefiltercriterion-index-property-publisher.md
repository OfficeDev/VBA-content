---
title: MailMergeFilterCriterion.Index Property (Publisher)
keywords: vbapb10.chm6815745
f1_keywords:
- vbapb10.chm6815745
ms.prod: publisher
api_name:
- Publisher.MailMergeFilterCriterion.Index
ms.assetid: e66e5afd-db28-cd00-9692-3b1a6d557198
ms.date: 06/08/2017
---


# MailMergeFilterCriterion.Index Property (Publisher)

Returns a  **Long** that represents the position of a particular item in a specified collection. .


## Syntax

 _expression_. **Index**

 _expression_A variable that represents a  **MailMergeFilterCriterion** object.


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


