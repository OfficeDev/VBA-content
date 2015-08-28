
# Report.Top Property (Access)

 **Last modified:** July 28, 2015

 **In this article**
 [Syntax](#sectionSection0)
 [Remarks](#sectionSection1)
 [Example](#sectionSection2)


You can use the  **Top** property to specify an object's location on a form or report. Read/write **Long**. .


## Syntax
<a name="sectionSection0"> </a>

 _expression_. **Top**

 _expression_A variable that represents a  **Report** object.


## Remarks
<a name="sectionSection1"> </a>

In Visual Basic, use a numeric expression to set the value of this property. Values are expressed in twips.

For reports, the  **Top** property setting is the amount the current section is offset from the top of the page. This property setting is expressed in twips. You can use this property to specify how far down the page you want a section to print in the section's **Format**event procedure.


## Example
<a name="sectionSection2"> </a>

The following example checks the  **Top** property setting for the current report. If the value is less than the minimum margin setting, the **NextRecord** and **PrintSection** properties are set to **False**. The section doesn't advance to the next record, and the next section isn't printed.


```
Sub Detail1_Format(Cancel As Integer, FormatCount As Integer) 
Const conTopMargin = 1880 
' Don't advance to next record or print next section 
' if Top property setting is less than 1880 twips. 
 If Me.Top < conTopMargin Then 
 Me.NextRecord = False 
 Me.PrintSection = False 
 End If 
End Sub
```


## See also
<a name="sectionSection2"> </a>


#### Concepts


 [Report Object](6f77c1b4-a9ce-7caa-204c-fe0755c6f9df.md)
#### Other resources


 [Report Object Members](73370a33-1ca0-da4d-9e36-88011bc2b93e.md)
