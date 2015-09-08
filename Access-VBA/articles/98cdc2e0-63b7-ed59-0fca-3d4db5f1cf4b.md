
# ToggleButton.Move Method (Access)

 **Last modified:** July 28, 2015

Moves the specified object to the coordinates specified by the argument values.

## Syntax

 _expression_. **Move**( **_Left_**,  **_Top_**,  **_Width_**,  **_Height_**)

 _expression_A variable that represents a  **ToggleButton** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|Left|Required| **Variant**|The screen position in twips for the left edge of the object relative to the left edge of the Microsoft Access window.|
|Top|Optional| **Variant**|The screen position in twips for the top edge of the object relative to the top edge of the Microsoft Access window.|
|Width|Optional| **Variant**|The desired width in twips of the object.|
|Height|Optional| **Variant**|The desired height in twips of the object.|

## Remarks

Only the Left argument is required. However, to specify any other arguments, you must specify all the arguments that precede it. For example, you cannot specifyWidth without specifyingLeft andTop. Any trailing arguments that are unspecified remain unchanged.

This method overrides the  **Moveable** property.

In Datasheet View or Print Preview, changes made using the  **Move** method are saved if the user explicitly saves the database, but Access does not prompt the user to save such changes.


## See also


#### Concepts


 [ToggleButton Object](1c20d809-d7db-e096-4328-ebb4d79e770e.md)
#### Other resources


 [ToggleButton Object Members](487101e7-c090-eb79-3671-5c9ce86cb6b0.md)
