
# OptionButton.ColumnWidth Property (Access)

 **Last modified:** July 28, 2015

 **In this article**
 [Syntax](#sectionSection0)
 [Remarks](#sectionSection1)
 [Example](#sectionSection2)


You can use the  **ColumnWidth** property to specify the width of a column in Datasheet view. Read/write **Integer**.


## Syntax
<a name="sectionSection0"> </a>

 _expression_. **ColumnWidth**

 _expression_A variable that represents an  **OptionButton** object.


## Remarks
<a name="sectionSection1"> </a>

In Visual Basic , the  **ColumnWidth** property setting is an **Integer** value that represents the column width in twips. You can specify a width or use one of the following predefined settings.



|**Setting**|**Description**|
|:-----|:-----|
|0|Hides the column.|
|-1|(Default) Sizes the column to the default width.|

 **Note**  The  **ColumnWidth** property applies to all fields in Datasheet view and to form controls when the form is in Datasheet view.

Setting this property to 0, or resizing the field to a zero width in Datasheet view, sets the field's  **ColumnHidden**property to  **True** (-1) and hides the field in Datasheet view.

Setting a field's  **ColumnHidden** property to **False** (0) restores the field's **ColumnWidth** property to the value it had before the field was hidden. For example, if the **ColumnWidth** property was -1 prior to the field being hidden by setting the property to 0, changing the field's **ColumnHidden** property to **False** resets the **ColumnWidth** to -1.

The  **ColumnWidth** property for a field isn't available when the field's **ColumnHidden** property is set to **True**.


## Example
<a name="sectionSection2"> </a>

This example takes effect in Datasheet view of the open Customers form. It sets the row height to 450 twips and sizes the column to fit the size of the visible text.


```
Forms![Customers].RowHeight = 450 
Forms![Customers]![Address].ColumnWidth = -2
```


## See also
<a name="sectionSection2"> </a>


#### Concepts


 [OptionButton Object](661ada74-d044-4a5c-2bdd-2dddfc2e79ab.md)
#### Other resources


 [OptionButton Object Members](5173d5c5-b898-97ee-a005-7f5a4d77efa1.md)
