
# AutoCaption.AutoInsert Property (Word)

 **Last modified:** July 28, 2015

 **True** if a caption is automatically added when the item is inserted into a document. Read/write **Boolean**.

## Syntax

 _expression_. **AutoInsert**

 _expression_A variable that represents an  ** [AutoCaption](895b5181-d36f-7f63-572a-c2d37c878e17.md)** object.


## Example

This example enables Word to add captions to tables automatically. Then the example collapses the selection to an insertion point, and inserts a table. A caption is automatically added to the new table.


```
AutoCaptions("Microsoft Word Table").AutoInsert = True 
Selection.Collapse Direction:=wdCollapseStart 
ActiveDocument.Tables.Add Range:=Selection.Range, _ 
 NumRows:=2, NumColumns:=2
```


## See also


#### Concepts


 [AutoCaption Object](895b5181-d36f-7f63-572a-c2d37c878e17.md)
#### Other resources


 [AutoCaption Object Members](48332cba-c2a5-a641-dc08-4cc2774ee5e6.md)
