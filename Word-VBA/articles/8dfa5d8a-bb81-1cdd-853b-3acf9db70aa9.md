
# Dialogs Object (Word)

 **Last modified:** July 28, 2015

A collection of  ** [Dialog](f90f6e6d-aaa0-c127-ab37-ca074144eff1.md)**objects in Word. Each  **Dialog** object represents a built-in Word dialog box.

## Remarks

Use the  ** [Dialogs](17acdfab-32d2-ddb8-04aa-692f9ffb20b8.md)** property to return the **Dialogs** collection. The following example displays the number of available built-in dialog boxes.


```
MsgBox Dialogs.Count
```

You cannot create a new built-in dialog box or add one to the  **Dialogs** collection. Use **Dialogs**(Index), where Index is the  **WdWordDialog** constant that identifies the dialog box, to return a single **Dialog** object. The following example displays the built-in **Open** dialog box.




```
dlgAnswer = Dialogs(wdDialogFileOpen).Show
```

For more information, see  [Displaying built-in Word dialog boxes](abe465f9-09a1-72ea-2e2d-9de14fc02434.md).


## See also


#### Concepts


 [Word Object Model Reference](be452561-b436-bb9b-6f94-3faa9a74a6fd.md)
#### Other resources


 [Dialogs Object Members](c1ab2260-007a-d276-787b-1cc91c35f93d.md)
