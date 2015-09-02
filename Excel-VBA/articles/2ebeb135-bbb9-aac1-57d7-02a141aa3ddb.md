
# Outline.ShowLevels Method (Excel)

 **Last modified:** July 28, 2015

 **In this article**
 [Syntax](#sectionSection0)
 [Remarks](#sectionSection1)
 [Example](#sectionSection2)


Displays the specified number of row and/or column levels of an outline.


## Syntax
<a name="sectionSection0"> </a>

 _expression_. **ShowLevels**( **_RowLevels_**,  **_ColumnLevels_**)

 _expression_A variable that represents an  **Outline** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|RowLevels|Optional| **Variant**|Specifies the number of row levels of an outline to display. If the outline has fewer levels than the number specified, Microsoft Excel displays all the levels. If this argument is 0 (zero) or is omitted, no action is taken on rows.|
|ColumnLevels|Optional| **Variant**|Specifies the number of column levels of an outline to display. If the outline has fewer levels than the number specified, Microsoft Excel displays all the levels. If this argument is 0 (zero) or is omitted, no action is taken on columns.|

### Return Value

Variant


## Remarks
<a name="sectionSection1"> </a>

You must specify at least one argument.


## Example
<a name="sectionSection2"> </a>

This example displays row levels one through three and column level one of the outline on Sheet1.


```
Worksheets("Sheet1").Outline _ 
 .ShowLevels rowLevels:=3, columnLevels:=1
```


## See also
<a name="sectionSection2"> </a>


#### Concepts


 [Outline Object](f5d50a8a-0dd9-638a-4374-5c648386a598.md)
#### Other resources


 [Outline Object Members](bf8e2103-d023-fc1f-90f2-960dff36e548.md)
