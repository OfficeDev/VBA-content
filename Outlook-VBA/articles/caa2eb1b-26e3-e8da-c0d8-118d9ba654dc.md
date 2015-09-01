
# View.Language Property (Outlook)

 **Last modified:** July 28, 2015

Returns or sets a  **String** value that represents the language setting for the object that defines the language used in the menu. Read/write.

## Syntax

 _expression_. **Language**

 _expression_A variable that represents a  **View** object.


## Remarks

The  **Language** property uses a **String** to represent an ISO language tag. For example, the string "EN-US" represents the ISO code for "United States - English."

If a valid language code is specified, the object will only be available in the  **View** menu for the specified language type. If no value is specified, the object item is available for all language types. The default value for this property is an empty string.


## See also


#### Concepts


 [View Object](41c8d149-9912-1685-4c8b-3c849cc6f1ed.md)
#### Other resources


 [View Object Members](ed3196c6-e779-64f7-db1d-e2fd22bb4688.md)
