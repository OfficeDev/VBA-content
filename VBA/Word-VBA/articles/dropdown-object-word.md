---
title: DropDown Object (Word)
keywords: vbawd10.chm2341
f1_keywords:
- vbawd10.chm2341
ms.prod: word
api_name:
- Word.DropDown
ms.assetid: 55233d61-d6d0-30f9-6825-ebbdbeb928b6
ms.date: 06/08/2017
---


# DropDown Object (Word)

Represents a drop-down form field that contains a list of items in a form.


## Remarks

Use  **FormFields** (index), where index is the index number or the bookmark name associated with the drop-down form field, to return a single **FormField** object. Use the **DropDown** property with the **FormField** object to return a **DropDown** object. The following example selects the first item in the drop-down form field named "DropDown" in the active document.


```
ActiveDocument.FormFields("DropDown1").DropDown.Value = 1
```

The index number represents the position of the form field in the  **[FormFields](formfields-object-word.md)** collection. The following example checks the type of the first form field in the active document. If it is a drop-down form field, the second item is selected.




```
If ActiveDocument.FormFields(1).Type = wdFieldFormDropDown Then 
 ActiveDocument.FormFields(1).DropDown.Value = 2 
End If
```

The following example determines whether form field represented by  _ffield_ is a valid drop-down form field before adding an item to it.




```
Set ffield = ActiveDocument.FormFields(1).DropDown 
If ffield.Valid = True Then 
 ffield.ListEntries.Add Name:="Hello" 
Else 
 MsgBox "First field is not a drop down" 
End If
```

Use the  **Add** method with the **FormFields** collection to add a drop-down form field. The following example adds a drop-down form field at the beginning of the active document and then adds items to the form field.




```
Set ffield = ActiveDocument.FormFields.Add( _ 
 Range:=ActiveDocument.Range(Start:=0, End:=0), _ 
 Type:=wdFieldFormDropDown) 
With ffield 
 .Name = "Colors" 
 With .DropDown.ListEntries 
 .Add Name:="Blue" 
 .Add Name:="Green" 
 .Add Name:="Red" 
 End With 
End With
```


## Properties



|**Name**|
|:-----|
|[Application](dropdown-application-property-word.md)|
|[Creator](dropdown-creator-property-word.md)|
|[Default](dropdown-default-property-word.md)|
|[ListEntries](dropdown-listentries-property-word.md)|
|[Parent](dropdown-parent-property-word.md)|
|[Valid](dropdown-valid-property-word.md)|
|[Value](dropdown-value-property-word.md)|

## See also


#### Other resources


[Word Object Model Reference](http://msdn.microsoft.com/library/be452561-b436-bb9b-6f94-3faa9a74a6fd%28Office.15%29.aspx)
