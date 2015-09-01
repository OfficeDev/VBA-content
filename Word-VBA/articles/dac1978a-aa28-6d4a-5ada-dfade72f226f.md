
# FormField.Creator Property (Word)

 **Last modified:** July 28, 2015

Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only  **Long**.

## Syntax

 _expression_. **Creator**

 _expression_Required. A variable that represents a  ** [FormField](c3c07344-06b2-fe86-6fcb-b9c63a991bcc.md)** object.


## Remarks

If the object was created in Microsoft Word, the  **Creator** property returns the hexadecimal number 4D535744, which represents the string "MSWD." This property was primarily designed to be used on the Macintosh, where each application has a four-character creator code. For example, Microsoft Word has the creator code MSWD. For additional information about this property, consult the language reference Help included with Microsoft Office Macintosh Edition.


## See also


#### Concepts


 [FormField Object](c3c07344-06b2-fe86-6fcb-b9c63a991bcc.md)
#### Other resources


 [FormField Object Members](e7d1b5d7-e1b3-b602-98c4-d0d4dc2288e5.md)
