
# Worksheet.Names Property (Excel)

 **Last modified:** July 28, 2015

 **In this article**
 [Syntax](#sectionSection0)
 [Remarks](#sectionSection1)
 [Example](#sectionSection2)


Returns a  ** [Names](ffecf89d-7bae-c470-8e37-608857a9de2a.md)** collection that represents all the worksheet-specific names (names defined with the "WorksheetName!" prefix). Read-only **Names** object.


## Syntax
<a name="sectionSection0"> </a>

 _expression_. **Names**

 _expression_A variable that represents a  **Worksheet** object.


## Remarks
<a name="sectionSection1"> </a>

Using this property without an object qualifier is equivalent to using  `ActiveWorkbook.Names`.


## Example
<a name="sectionSection2"> </a>

This example defines the name "myName" for cell A1 on Sheet1.


```
ActiveWorkbook.Names.Add Name:="myName", RefersToR1C1:= _ 
 "=Sheet1!R1C1"
```


## See also
<a name="sectionSection2"> </a>


#### Concepts


 [Worksheet Object](182b705e-854a-81cc-a4b0-59b942de55ae.md)
#### Other resources


 [Worksheet Object Members](f8c1afea-1a1c-f5e4-37e3-52c434c8c157.md)
