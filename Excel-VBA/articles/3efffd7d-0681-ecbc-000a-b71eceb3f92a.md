
# Workbook.BuiltinDocumentProperties Property (Excel)

 **Last modified:** July 28, 2015

 **In this article**
 [Syntax](#sectionSection0)
 [Remarks](#sectionSection1)
 [Example](#sectionSection2)


Returns a  ** [DocumentProperties](http://msdn.microsoft.com/library/90d42786-7d9a-b604-dbdf-88db41cbe69b%28Office.15%29.aspx)** collection that represents all the built-in document properties for the specified workbook. Read-only.


## Syntax
<a name="sectionSection0"> </a>

 _expression_. **BuiltinDocumentProperties**

 _expression_A variable that represents a  **Workbook** object.


## Remarks
<a name="sectionSection1"> </a>

This property returns the entire collection of built-in document properties. Use the  **Item** method to return a single member of the collection (a ** [DocumentProperty](http://msdn.microsoft.com/library/dd54ca3c-e0e2-4816-539a-17c5b4a928b1%28Office.15%29.aspx)** object) by specifying either the name of the property or the collection index (as a number).

You can refer to document properties either by index value or by name. The following list shows the available built-in document property names:



|Title Subject Author Keywords Comments Template Last Author Revision Number Application Name Last Print Date|Creation Date Last Save Time Total Editing Time Number of Pages Number of Words Number of Characters Security Category Format Manager|Company Number of Bytes Number of Lines Number of Paragraphs Number of Slides Number of Notes Number of Hidden Slides Number of Multimedia Clips Hyperlink Base Number of Characters (with spaces)|
Container applications aren't required to define values for every built-in document property. If Microsoft Excel doesn't define a value for one of the built-in document properties, reading the  **Value** property for that document property causes an error.

Because the  **Item** method is the default method for the **DocumentProperties** collection, the following statements are identical:




```
BuiltinDocumentProperties.Item(1) 
BuiltinDocumentProperties(1)
```

Use the  ** [CustomDocumentProperties](8470adbb-5b10-96ba-71f7-c667c33b6707.md)** property to return the collection of custom document properties.


## Example
<a name="sectionSection2"> </a>

This example displays the names of the built-in document properties as a list on worksheet one.


```
rw = 1 
Worksheets(1).Activate 
For Each p In ActiveWorkbook.BuiltinDocumentProperties 
    Cells(rw, 1).Value = p.Name 
    rw = rw + 1 
Next
```


## See also
<a name="sectionSection2"> </a>


#### Concepts


 [Workbook Object](8c00aa60-c974-eed3-0812-3c9625eb0d4c.md)
#### Other resources


 [Workbook Object Members](dce102a3-25de-3ff4-2ce5-bc56e08baca7.md)
