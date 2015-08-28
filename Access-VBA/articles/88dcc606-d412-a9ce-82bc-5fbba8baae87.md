
# Range.Errors Property (Excel)

 **Last modified:** July 28, 2015

 **In this article**
 [Syntax](#sectionSection0)
 [Remarks](#sectionSection1)
 [Example](#sectionSection2)


Allows the user to to access error checking options.


## Syntax
<a name="sectionSection0"> </a>

 _expression_. **Errors**

 _expression_A variable that represents a  **Range** object.


## Remarks
<a name="sectionSection1"> </a>

Reference the  ** [Errors](d2b50bbf-2685-fc5f-74c5-fa8bb9955f2a.md)**object to view a list of index values associated with error checking options.


## Example
<a name="sectionSection2"> </a>

In this example, a number written as text is placed in cell A1. Microsoft Excel then determines if the number is written as text in cell A1 and notifies the user accordingly.


```
Sub CheckForErrors() 
 
 Range("A1").Formula = "'12" 
 
 If Range("A1").Errors.Item(xlNumberAsText).Value = True Then 
 MsgBox "The number is written as text." 
 Else 
 MsgBox "The number is not written as text." 
 End If 
 
End Sub
```


## See also
<a name="sectionSection2"> </a>


#### Concepts


 [Range Object](b8207778-0dcc-4570-1234-f130532cc8cd.md)
#### Other resources


 [Range Object Members](4336bf81-1e63-7e44-1792-baf366a027a7.md)
