
# Shapes.AddLabel Method (Excel)

 **Last modified:** July 28, 2015

Creates a label. Returns a  ** [Shape](8f01fcd1-b7d9-5216-2de5-40fb6648a403.md)** object that represents the new label.

## Syntax

 _expression_. **AddLabel**( **_Orientation_**,  **_Left_**,  **_Top_**,  **_Width_**,  **_Height_**)

 _expression_A variable that represents a  **Shapes** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|Orientation|Required| ** [MsoTextOrientation](http://msdn.microsoft.com/library/7e8d0e06-14dd-3ea1-a2e4-50375919517f%28Office.15%29.aspx)**|The text orientation within the label.|
|Left|Required| **Single**|The position (in points) of the upper-left corner of the label relative to the upper-left corner of the document.|
|Top|Required| **Single**|The position (in points) of the upper-left corner of the label relative to the top corner of the document.|
|Width|Required| **Single**|The width of the label, in points.|
|Height|Required| **Single**|The height of the label, in points.|

### Return Value

Shape


## Example

This example adds a vertical label that contains the text "Test Label" to  `myDocument`.


```
Set myDocument = Worksheets(1) 
myDocument.Shapes.AddLabel(msoTextOrientationVertical, _ 
    100, 100, 60, 150) _ 
    .TextFrame.Characters.Text = "Test Label"
```


## See also


#### Concepts


 [Shapes Object](f9c6548c-d028-1b70-a11c-c4b45ff19177.md)
#### Other resources


 [Shapes Object Members](f5d0be42-46cc-2916-8953-401e50a5cef7.md)
