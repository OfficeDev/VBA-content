
# Table.AutoFitBehavior Method (Word)

 **Last modified:** July 28, 2015

 **In this article**
 [Syntax](#sectionSection0)
 [Remarks](#sectionSection1)
 [Example](#sectionSection2)


Determines how Microsoft Word resizes a table when the AutoFit feature is used.


## Syntax
<a name="sectionSection0"> </a>

 _expression_. **AutoFitBehavior**( **_Behavior_**)

 _expression_Required. A variable that represents a  ** [Table](996b58dd-ebc6-ee30-5bfe-c5e51a0f71d6.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|Behavior|Required| **WdAutoFitBehavior**|How Word resizes the specified table with the AutoFit feature is used.|

## Remarks
<a name="sectionSection1"> </a>

Word can resize the table based on the content of the table cells or the width of the document window. You can also use this method to turn off AutoFit so that the table size is fixed, regardless of cell contents or window width.

Setting the  **AutoFitBehavior** property to **wdAutoFitContent** or **wdAutoFitWindow** sets the **AllowAutoFit**property to  **True** if it is currently **False**. Likewise, setting the  **AutoFitBehavior** property to **wdAutoFitFixed** sets the **AllowAutoFit** property to **False** if it is currently **True**.


## Example
<a name="sectionSection2"> </a>

This example sets the AutoFit behavior for the first table in the active document to automatically resize based on the width of the document window.


```
ActiveDocument.Tables(1).AutoFitBehavior _ 
 wdAutoFitWindow
```


## See also
<a name="sectionSection2"> </a>


#### Concepts


 [Table Object](996b58dd-ebc6-ee30-5bfe-c5e51a0f71d6.md)
#### Other resources


 [Table Object Members](5367ee92-b5a3-92c7-787b-46a302586a0d.md)
