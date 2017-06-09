---
title: FormFields Object (Word)
ms.prod: word
ms.assetid: a44a0f57-123b-cade-e306-ba6dc179b619
ms.date: 06/08/2017
---


# FormFields Object (Word)

A collection of  **FormField** objects that represent all the form fields in a selection, range, or document.


## Remarks

Use the  **FormFields** property to return the **FormFields** collection. The following example counts the number of text box form fields in the active document.


```
For Each aField In ActiveDocument.FormFields 
 If aField.Type = wdFieldFormTextInput Then count = count + 1 
Next aField 
MsgBox "There are " &amp; count &amp; " text boxes in this document"
```

Use the  **Add** method with the **FormFields** object to add a form field. The following example adds a check box at the beginning of the active document and then selects the check box.




```
Set ffield = ActiveDocument.FormFields.Add( _ 
 Range:=ActiveDocument.Range(Start:=0,End:=0), _ 
 Type:=wdFieldFormCheckBox) 
ffield.CheckBox.Value = True
```

Use  **FormFields** (Index), where Index is a bookmark name or index number, to return a single **[FormField](formfield-object-word.md)** object. The following example sets the result of the Text1 form field to "Don Funk."




```
ActiveDocument.FormFields("Text1").Result = "Don Funk"
```

The index number represents the position of the form field in the selection, range, or document. The following example displays the name of the first form field in the selection.




```
If Selection.FormFields.Count >= 1 Then 
 MsgBox Selection.FormFields(1).Name 
End If
```


## Methods



|**Name**|
|:-----|
|[Add](formfields-add-method-word.md)|
|[Item](formfields-item-method-word.md)|

## Properties



|**Name**|
|:-----|
|[Application](formfields-application-property-word.md)|
|[Count](formfields-count-property-word.md)|
|[Creator](formfields-creator-property-word.md)|
|[Parent](formfields-parent-property-word.md)|
|[Shaded](formfields-shaded-property-word.md)|

## See also


#### Other resources


[Word Object Model Reference](http://msdn.microsoft.com/library/be452561-b436-bb9b-6f94-3faa9a74a6fd%28Office.15%29.aspx)
