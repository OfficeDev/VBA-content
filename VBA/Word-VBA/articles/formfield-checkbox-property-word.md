---
title: FormField.CheckBox Property (Word)
keywords: vbawd10.chm153616396
f1_keywords:
- vbawd10.chm153616396
ms.prod: word
api_name:
- Word.FormField.CheckBox
ms.assetid: 6843d3e0-8f34-422f-403e-3bab806dc6be
ms.date: 06/08/2017
---


# FormField.CheckBox Property (Word)

Returns a  **[CheckBox](checkbox-object-word.md)** object that represents a check box form field. Read-only.


## Syntax

 _expression_ . **CheckBox**

 _expression_ A variable that represents a **[FormField](formfield-object-word.md)** object.


## Remarks

If the  **CheckBox** property is applied to a **FormField** object that isn't a check box form field, the property won't fail, but the **Valid** property for the returned object will be **False** .


## Example

This example clears the check box named "Blue."


```vb
ActiveDocument.FormFields("Blue").CheckBox.Value = False
```

This example compares the current value with the default value of the check box named "Check1." If the values are equal, the blnSame variable is set to True.




```vb
Dim ffTemp As FormField 
Dim blnSame As Boolean 
 
Set ffTemp = ActiveDocument.FormFields("Check1").CheckBox 
If ffTemp.Default = ffTemp.Value Then 
 blnSame = True 
Else 
 blnSame = False 
End If
```


## See also


#### Concepts


[FormField Object](formfield-object-word.md)

