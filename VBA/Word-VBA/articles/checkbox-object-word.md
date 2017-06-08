---
title: CheckBox Object (Word)
keywords: vbawd10.chm2342
f1_keywords:
- vbawd10.chm2342
ms.prod: word
api_name:
- Word.CheckBox
ms.assetid: e72b57b7-0328-9e78-94ca-ab7fb3c64afb
ms.date: 06/08/2017
---


# CheckBox Object (Word)

Represents a single check box form field.


## Remarks

Use  **FormFields** (Index), where Index is index number or the bookmark name associated with the check box, to return a single **[FormField](formfield-object-word.md)** object. Use the **[CheckBox](formfield-checkbox-property-word.md)** property with the **FormField** object to return a **CheckBox** object. The following example selects the check box form field named "Check1" in the active document.


```
ActiveDocument.FormFields("Check1").CheckBox.Value = True
```

The index number represents the position of the form field in the  **[FormFields](formfields-object-word.md)** collection. The following example checks the type of the first form field; if it is a check box, the check box is selected.




```
If ActiveDocument.FormFields(1).Type = wdFieldFormCheckBox Then 
 ActiveDocument.FormFields(1).CheckBox.Value = True 
End If
```

The following example determines whether the  _ffield_ object is valid before changing the check box size to 14 points.




```
Set ffield = ActiveDocument.FormFields(1).CheckBox 
If ffield.Valid = True Then 
 ffield.AutoSize = False 
 ffield.Size = 14 
Else 
 MsgBox "First field is not a check box" 
End If
```

Use the  **Add** method with the **FormFields** object to add a check box form field. The following example adds a check box at the beginning of the active document, sets the name to "Color", and then selects the check box.




```
With ActiveDocument.FormFields.Add(Range:=ActiveDocument.Range _ 
 (Start:=0,End:=0), Type:=wdFieldFormCheckBox) 
 .Name = "Color" 
 .CheckBox.Value = True 
End With
```


## Properties



|**Name**|
|:-----|
|[Application](checkbox-application-property-word.md)|
|[AutoSize](checkbox-autosize-property-word.md)|
|[Creator](checkbox-creator-property-word.md)|
|[Default](checkbox-default-property-word.md)|
|[Parent](checkbox-parent-property-word.md)|
|[Size](checkbox-size-property-word.md)|
|[Valid](checkbox-valid-property-word.md)|
|[Value](checkbox-value-property-word.md)|

## See also


#### Other resources


[Word Object Model Reference](http://msdn.microsoft.com/library/be452561-b436-bb9b-6f94-3faa9a74a6fd%28Office.15%29.aspx)
