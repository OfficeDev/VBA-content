
# Application.Quit Event (Outlook)

Occurs when Microsoft Outlook begins to close. 


## Syntax

 _expression_ . **Quit**

 _expression_ An expression that returns an **Application** object.


## Remarks

This event is not available in Microsoft Visual Basic Scripting Edition (VBScript).


## Example

This Microsoft Visual Basic for Applications (VBA) example displays a farewell message when Outlook exits. The sample code must be placed in a class module.


```vb
Private Sub Application_Quit() 
 
 MsgBox "Goodbye, " &; Application.GetNamespace("MAPI").CurrentUser 
 
End Sub
```


## See also


#### Concepts


[Application Object](797003e7-ecd1-eccb-eaaf-32d6ddde8348.md)
#### Other resources


[Application Object Members](3519c89c-2353-85ee-7ddc-62e5dd85a8e7.md)
