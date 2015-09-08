
# OptionButton.Move Method (Access)

 **Last modified:** July 28, 2015

Moves the specified object to the coordinates specified by the argument values.

## Syntax

 _expression_. **Move**( **_Left_**,  **_Top_**,  **_Width_**,  **_Height_**)

 _expression_A variable that represents an  **OptionButton** object.


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


 [OptionButton Object](661ada74-d044-4a5c-2bdd-2dddfc2e79ab.md)
#### Other resources


 [OptionButton Object Members](5173d5c5-b898-97ee-a005-7f5a4d77efa1.md)
