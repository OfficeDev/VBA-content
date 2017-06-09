---
title: CustomLabels Object (Word)
ms.prod: word
ms.assetid: 407e75b5-4116-fdc7-f0c1-dfd3809cdb41
ms.date: 06/08/2017
---


# CustomLabels Object (Word)

A collection of  **CustomLabel** objects available in the **Label Options** dialog box. This collection includes custom labels of all printer types (dot-matrix, laser, and ink-jet printers).


## Remarks

Use the  **CustomLabels** property to return the **CustomLabels** collection. The following example displays the number of available custom labels.


```vb
MsgBox Application.MailingLabel.CustomLabels.Count
```

Use the  **[Add](customlabels-add-method-word.md)** method to create a custom label. The following example adds a custom mailing label named "My Label" and sets the page size.




```vb
Set ML = _ 
 Application.MailingLabel.CustomLabels.Add(Name:="My Labels", _ 
 DotMatrix:=False) 
ML.PageSize = wdCustomLabelA4
```

Use  **[CustomLabels](mailinglabel-customlabels-property-word.md)** (Index), where Index is the custom label name or index number, to return a single **[CustomLabel](customlabel-object-word.md)** object. The following example creates a new document with an existing custom label layout named "My Labels."




```vb
Set ML = Application.MailingLabel 
If ML.CustomLabels("My Labels").Valid = True Then 
 ML.CreateNewDocument Name:="My Labels" 
Else 
 MsgBox "The My Labels custom label is not available" 
End If
```

The index number represents the position of the custom mailing label in the  **CustomLabels** collection. The following example displays the name of the first custom mailing label.




```vb
If Application.MailingLabel.CustomLabels.Count >= 1 Then 
 MsgBox Application.MailingLabel.CustomLabels(1).Name 
End If
```


## See also


#### Other resources



[Word Object Model Reference](http://msdn.microsoft.com/library/be452561-b436-bb9b-6f94-3faa9a74a6fd%28Office.15%29.aspx)

