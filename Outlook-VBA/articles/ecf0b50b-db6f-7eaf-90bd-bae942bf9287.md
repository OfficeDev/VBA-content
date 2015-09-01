
# Application.Quit Event (Outlook)

 **Last modified:** July 28, 2015

 **In this article**
 [Syntax](#sectionSection0)
 [Remarks](#sectionSection1)
 [Example](#sectionSection2)


Occurs when Microsoft Outlook begins to close. 


## Syntax
<a name="sectionSection0"> </a>

 _expression_. **Quit**

 _expression_An expression that returns an  **Application** object.


## Remarks
<a name="sectionSection1"> </a>

This event is not available in Microsoft Visual Basic Scripting Edition (VBScript).


## Example
<a name="sectionSection2"> </a>

This Microsoft Visual Basic for Applications (VBA) example displays a farewell message when Outlook exits. The sample code must be placed in a class module.


```
Private Sub Application_Quit() 
 
 MsgBox "Goodbye, " &amp; Application.GetNamespace("MAPI").CurrentUser 
 
End Sub
```


## See also
<a name="sectionSection2"> </a>


#### Concepts


 [Application Object](797003e7-ecd1-eccb-eaaf-32d6ddde8348.md)
#### Other resources


 [Application Object Members](3519c89c-2353-85ee-7ddc-62e5dd85a8e7.md)
