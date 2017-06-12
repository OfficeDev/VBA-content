---
title: Referencing Controls on an Outlook Form
keywords: olfm10.chm3077115
f1_keywords:
- olfm10.chm3077115
ms.prod: outlook
ms.assetid: 1393bd23-de16-4a59-e656-f0fcc6583a3e
ms.date: 06/08/2017
---


# Referencing Controls on an Outlook Form

If you need to refer to a control on an Outlook form within a procedure, you must also reference the inspector, page, and controls collection that contains the control, even if you are referencing the control with its own event procedure. The following example shows how to change the caption of a command button when it's clicked. To test this example, in design mode create a command button with the default name CommandButton1 on the page P.2.


```vb
Sub CommandButton1_Click 
 Set myButton = Item.GetInspector.ModifiedFormPages("P.2")_ 
 .Controls("CommandButton1") 
 myButton.Caption = "New Caption" 
End Sub
```


