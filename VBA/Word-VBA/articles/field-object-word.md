---
title: Field Object (Word)
keywords: vbawd10.chm2351
f1_keywords:
- vbawd10.chm2351
ms.prod: word
api_name:
- Word.Field
ms.assetid: 75139aa4-89f4-2ffb-b964-8dc805b9a32b
ms.date: 06/08/2017
---


# Field Object (Word)

Represents a field. The  **Field** object is a member of the **Fields** collection. The **[Fields](fields-object-word.md)** collection represents the fields in a selection, range, or document.


## Remarks

Use  **Fields** (Index), where Index is the index number, to return a single **Field** object. The index number represents the position of the field in the selection, range, or document. The following example displays the field code and the result of the first field in the active document.


```vb
If ActiveDocument.Fields.Count >= 1 Then 
 MsgBox "Code = " &; ActiveDocument.Fields(1).Code &; vbCr _ 
 &; "Result = " &; ActiveDocument.Fields(1).Result &; vbCr 
End If
```

Use the  **Add** method to add a field to the **[Fields](fields-object-word.md)** collection. The following example inserts a DATE field at the beginning of the selection and then displays the result.




```vb
Selection.Collapse Direction:=wdCollapseStart 
Set myField = ActiveDocument.Fields.Add(Range:=Selection.Range, _ 
 Type:=wdFieldDate) 
MsgBox myField.Result
```


 **Note**  The  **wdFieldDate** constant is part of the **[WdFieldType](wdfieldtype-enumeration-word.md)** group of constants, which includes all the various field types.


## See also


#### Other resources


[Word Object Model Reference](http://msdn.microsoft.com/library/be452561-b436-bb9b-6f94-3faa9a74a6fd%28Office.15%29.aspx)


