
# Worksheet.Names Property (Excel)

Returns a  **[Names](ffecf89d-7bae-c470-8e37-608857a9de2a.md)** collection that represents all the worksheet-specific names (names defined with the "WorksheetName!" prefix). Read-only **Names** object.


## Syntax

 _expression_ . **Names**

 _expression_ A variable that represents a **Worksheet** object.


## Remarks

Using this property without an object qualifier is equivalent to using  `ActiveWorkbook.Names`.


## Example

This example defines the name "myName" for cell A1 on Sheet1.


```vb
ActiveWorkbook.Names.Add Name:="myName", RefersToR1C1:= _ 
 "=Sheet1!R1C1"
```


## See also


#### Concepts


[Worksheet Object](182b705e-854a-81cc-a4b0-59b942de55ae.md)
#### Other resources


[Worksheet Object Members](f8c1afea-1a1c-f5e4-37e3-52c434c8c157.md)
