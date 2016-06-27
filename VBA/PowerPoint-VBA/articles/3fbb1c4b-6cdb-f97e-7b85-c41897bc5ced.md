
# Table.TableDirection Property (PowerPoint)

Returns or sets the direction in which the table cells are ordered. Read/write.


## Syntax

 _expression_. **TableDirection**

 _expression_ A variable that represents a **Table** object.


### Return Value

PpDirection


## Remarks

The default value of the  **TableDirection** property is **ppDirectionLefttToRight**, unless the **[LanguageSettings](9603b5ed-2143-10f7-399b-2757b71c0525.md)** property or the **[DefaultLanguageID](8568c96c-b997-6a92-e93b-0f3d091383e2.md)** property is set to a right-to-left language, in which case the default value is **ppDirectionRightToLeft**.

The value of the  **TableDirection** property can be one of these **PpDirection** constants.


||
|:-----|
|**ppDirectionLeftToRight**|
|**ppDirectionMixed**|
|**ppDirectionRightToLeft**|
When you are using the  **TextDirection** property, The **ppDirectionMixed** constant may be returned.


## Example

This example sets the direction in which cells in the selected table are ordered to left to right (first column is the leftmost column).


```vb
ActiveWindow.Selection.ShapeRange.Table.TableDirection = _
    ppDirectionLeftToRight
```


## See also


#### Concepts


[Table Object](ebbbca9f-4591-10ce-3c74-33b46a3b7cdf.md)
#### Other resources


[Table Object Members](97f64cfc-1762-c935-6714-b5c5b5a6cc3c.md)
