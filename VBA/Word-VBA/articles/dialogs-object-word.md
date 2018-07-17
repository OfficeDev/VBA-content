---
title: Dialogs Object (Word)
ms.prod: word
ms.assetid: 8dfa5d8a-bb81-1cdd-853b-3acf9db70aa9
ms.date: 06/08/2017
---


# Dialogs Object (Word)

A collection of  **[Dialog](dialog-object-word.md)** objects in Word. Each **Dialog** object represents a built-in Word dialog box.


## Remarks

Use the  **[Dialogs](application-dialogs-property-word.md)** property to return the **Dialogs** collection. The following example displays the number of available built-in dialog boxes.


```vb
MsgBox Dialogs.Count
```

You cannot create a new built-in dialog box or add one to the  **Dialogs** collection. Use **Dialogs** (Index), where Index is the **[WdWordDialog](wdworddialog-enumeration-word.md)** constant that identifies the dialog box, to return a single **Dialog** object. The following example displays the built-in **Open** dialog box.




```
dlgAnswer = Dialogs(wdDialogFileOpen).Show
```

For more information, see [Displaying built-in Word dialog boxes](displaying-built-in-word-dialog-boxes.md).


## See also


#### Other resources


[Word Object Model Reference](http://msdn.microsoft.com/library/be452561-b436-bb9b-6f94-3faa9a74a6fd%28Office.15%29.aspx)


