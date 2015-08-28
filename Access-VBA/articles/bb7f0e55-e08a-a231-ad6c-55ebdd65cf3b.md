
# OptionButton.Left Property (Access)

 **Last modified:** July 28, 2015

 **In this article**
 [Syntax](#sectionSection0)
 [Remarks](#sectionSection1)
 [Example](#sectionSection2)


You can use the  **Left** property to specify an object's location on a form or report. Read/write **Integer**.


## Syntax
<a name="sectionSection0"> </a>

 _expression_. **Left**

 _expression_A variable that represents an  **OptionButton** object.


## Remarks
<a name="sectionSection1"> </a>

In Visual Basic, use a numeric expression to set the value of this property. Values are expressed in twips

For reports, you can set these properties only by using a macro or event procedure in Visual Basic while the report is in Print Preview or being printed.

For reports, the  **Left** property setting is the amount the current section is offset from the left of the page. This property is expressed in twips. You can use this property to specify how far down the page you want a section to print in the section's **Format**event procedure.


## Example
<a name="sectionSection2"> </a>

The following example checks the  **Left** property setting for the current report. If the value is less than the minimum margin setting, the **NextRecord** and **PrintSection** properties are set to **False** (0). The section doesn't advance to the next record, and the next section isn't printed.


```
Sub Detail1_Format(Cancel As Integer, FormatCount As Integer) 
 
 Const conLeftMargin = 1880 
 
 ' Don't advance to next record or print next section 
 ' if Left property setting is less than 1880 twips. 
 If Me.Left < conLeftMargin Then 
 Me.NextRecord = False 
 Me.PrintSection = False 
 End If 
 
End Sub
```


## See also
<a name="sectionSection2"> </a>


#### Concepts


 [OptionButton Object](661ada74-d044-4a5c-2bdd-2dddfc2e79ab.md)
#### Other resources


 [OptionButton Object Members](5173d5c5-b898-97ee-a005-7f5a4d77efa1.md)
