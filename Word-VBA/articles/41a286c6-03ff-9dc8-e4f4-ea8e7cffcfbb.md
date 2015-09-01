
# Language.Creator Property (Word)

 **Last modified:** July 28, 2015

Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only  **Long**.

## Syntax

 _expression_. **Creator**

 _expression_Required. A variable that represents a  ** [Language](0acc4a42-b4c2-a415-0e38-a049b085dc86.md)** object.


## Remarks

If the object was created in Microsoft Word, the  **Creator** property returns the hexadecimal number 4D535744, which represents the string "MSWD." This property was primarily designed to be used on the Macintosh, where each application has a four-character creator code. For example, Microsoft Word has the creator code MSWD. For additional information about this property, consult the language reference Help included with Microsoft Office Macintosh Edition.


## See also


#### Concepts


 [Language Object](0acc4a42-b4c2-a415-0e38-a049b085dc86.md)
#### Other resources


 [Language Object Members](71b8c7ea-bb8f-3fa7-73f7-f99485ab5d4a.md)
