
# ShapeNode.Creator Property (Word)

 **Last modified:** July 28, 2015

Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only  **Long**.

## Syntax

 _expression_. **Creator**

 _expression_Required. A variable that represents a  ** [ShapeNode](d5afb71a-a218-57f3-87f0-171094ba6610.md)** object.


## Remarks

If the object was created in Microsoft Word, the  **Creator** property returns the hexadecimal number 4D535744, which represents the string "MSWD." This property was primarily designed to be used on the Macintosh, where each application has a four-character creator code. For example, Microsoft Word has the creator code MSWD. For additional information about this property, consult the language reference Help included with Microsoft Office Macintosh Edition.


## See also


#### Concepts


 [ShapeNode Object](d5afb71a-a218-57f3-87f0-171094ba6610.md)
#### Other resources


 [ShapeNode Object Members](55803c23-5f6e-aa8c-6e9f-6d350ec71f5e.md)
