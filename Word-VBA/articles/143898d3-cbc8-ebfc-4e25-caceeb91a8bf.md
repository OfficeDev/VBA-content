
# Axes.Item Method (Word)

 **Last modified:** July 28, 2015

Returns a single  ** [Axis](3a7ad7d8-d397-a79a-eb6a-a5f0822cbe5d.md)** object from an **Axes** collection.

## Syntax

 _expression_. **Item**( **_Type_**,  **_AxisGroup_**)

 _expression_A variable that represents an  ** [Axes](57261ca9-7fd6-ba99-19bd-5df8e940f714.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|Type|Required| ** [XlAxisType](f02ed77e-8315-f318-ded2-751bc72d19fc.md)**|One of the enumeration values that specifies the axis type.|
|AxisGroup|Optional| ** [XlAxisGroup](ed3ff1ce-28de-165d-bbfa-f3d770f32522.md)**|One of the enumeration values that specifies the axis.|

## Example

The following example sets the title text for the category axis for the first chart in the active document.


```
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 With .Chart.Axes.Item(xlCategory) 
 .HasTitle = True 
 .AxisTitle.Caption = "1994" 
 End With 
 End If 
End With
```


## See also


#### Concepts


 [Axes Object](57261ca9-7fd6-ba99-19bd-5df8e940f714.md)
#### Other resources


 [Axes Object Members](dfbf9171-9d13-3555-4bb2-47d7fb98928a.md)
