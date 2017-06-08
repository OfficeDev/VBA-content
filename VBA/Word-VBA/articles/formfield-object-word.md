---
title: FormField Object (Word)
keywords: vbawd10.chm2344
f1_keywords:
- vbawd10.chm2344
ms.prod: word
api_name:
- Word.FormField
ms.assetid: c3c07344-06b2-fe86-6fcb-b9c63a991bcc
ms.date: 06/08/2017
---


# FormField Object (Word)

Represents a single form field. The  **FormField** object is a member of the **FormFields** collection.


## Remarks

Use  **FormFields** (index), where index is a bookmark name or index number, to return a single **FormField** object. The following example sets the result of the Text1 form field to "Don Funk."


```vb
ActiveDocument.FormFields("Text1").Result = "Don Funk"
```

The index number represents the position of the form field in the selection, range, or document. The following example displays the name of the first form field in the selection.




```vb
If Selection.FormFields.Count >= 1 Then 
 MsgBox Selection.FormFields(1).Name 
End If
```

Use the  **Add** method with the **[FormFields](formfields-object-word.md)** object to add a form field. The following example adds a check box at the beginning of the active document and then selects the check box.




```vb
Set ffield = ActiveDocument.FormFields.Add( _ 
 Range:=ActiveDocument.Range(Start:=0, End:=0), _ 
 Type:=wdFieldFormCheckBox) 
ffield.CheckBox.Value = True
```

Use the  **CheckBox** , **DropDown** , and **TextInput** properties with the **FormField** object to return the **CheckDown** , **DropDown** , and **TextInput** objects. The following example selects the check box named "Check1."




```vb
ActiveDocument.FormFields("Check1").CheckBox.Value = True
```


## See also


#### Other resources



[Word Object Model Reference](http://msdn.microsoft.com/library/be452561-b436-bb9b-6f94-3faa9a74a6fd%28Office.15%29.aspx)

