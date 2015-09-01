
# TwoInitialCapsExceptions.Creator Property (Word)

 **Last modified:** July 28, 2015

Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only  **Long**.

## Syntax

 _expression_. **Creator**

 _expression_Required. A variable that represents a  ** [TwoInitialCapsExceptions](21af2d69-8d76-026d-2002-8d69b4ab8aef.md)** collection.


## Remarks

If the object was created in Microsoft Word, the  **Creator** property returns the hexadecimal number 4D535744, which represents the string "MSWD." This property was primarily designed to be used on the Macintosh, where each application has a four-character creator code. For example, Microsoft Word has the creator code MSWD. For additional information about this property, consult the language reference Help included with Microsoft Office Macintosh Edition.


## See also


#### Concepts


 [TwoInitialCapsExceptions Collection Object](21af2d69-8d76-026d-2002-8d69b4ab8aef.md)
#### Other resources


 [TwoInitialCapsExceptions Object Members](05f0a660-a906-3d20-0190-99b23153fe73.md)
