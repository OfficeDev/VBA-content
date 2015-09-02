
# FullSeriesCollection.Application Property (Excel)

 **Last modified:** July 28, 2015

 **In this article**
 [Syntax](#sectionSection0)
 [Example](#sectionSection1)
 [Property value](#sectionSection2)


Returns an  ** [Application](19b73597-5cf9-4f56-8227-b5211f657f6f.md)** object that represents the Microsoft Excel application. Read-only.


## Syntax
<a name="sectionSection0"> </a>

 _expression_. **Application**

 _expression_A variable that represents a  [FullSeriesCollection Object (Excel)](5d7b7e7c-0a74-307b-84f9-56143ceba464.md) object.


## Example
<a name="sectionSection1"> </a>

This example displays a message about the application that created  `myObject`.


```
Set myObject = ActiveWorkbook 
If myObject.Application.Value = "Microsoft Excel" Then 
 MsgBox "This is an Excel Application object." 
Else 
 MsgBox "This is not an Excel Application object." 
End If
```


## Property value
<a name="sectionSection2"> </a>

 **APPLICATION**


## See also
<a name="sectionSection2"> </a>


#### Other resources


 [FullSeriesCollection Object Members](18060b3a-f25c-fa99-d3f3-dd59f7928465.md)
 [FullSeriesCollection Object](5d7b7e7c-0a74-307b-84f9-56143ceba464.md)
