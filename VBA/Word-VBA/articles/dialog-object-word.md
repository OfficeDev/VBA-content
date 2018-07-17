---
title: Dialog Object (Word)
keywords: vbawd10.chm2488
f1_keywords:
- vbawd10.chm2488
ms.prod: word
api_name:
- Word.Dialog
ms.assetid: f90f6e6d-aaa0-c127-ab37-ca074144eff1
ms.date: 06/08/2017
---


# Dialog Object (Word)

Represents a built-in dialog box. The  **Dialog** object is a member of the **[Dialogs](dialogs-object-word.md)** collection. The **Dialogs** collection contains all the built-in dialog boxes in Word. You cannot create a new built-in dialog box or add one to the **Dialogs** collection.


## Remarks

Use  **Dialogs** (Index), where Index is a **WdWordDialog** constant that identifies the dialog box, to return a single **Dialog** object. The following example displays and carries out the actions taken in the built-in **Open** dialog box.


```
dlgAnswer = Dialogs(wdDialogFileOpen).Show
```

The  **WdWordDialog** constants are formed from the prefix "wdDialog" followed by the name of the menu and the dialog box. For example, the constant for the **Page Setup** dialog box is **wdDialogFilePageSetup** , and the constant for the **New** dialog box is **wdDialogFileNew** .

For more information about working with built-in Word dialog boxes, see [Displaying built-in Word dialog boxes](http://msdn.microsoft.com/library/abe465f9-09a1-72ea-2e2d-9de14fc02434%28Office.15%29.aspx).


## See also


#### Other resources


[Word Object Model Reference](http://msdn.microsoft.com/library/be452561-b436-bb9b-6f94-3faa9a74a6fd%28Office.15%29.aspx)


