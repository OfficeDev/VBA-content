
# CustomLabels.Creator Property (Word)

 **Last modified:** July 28, 2015

Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only  **Long**.

## Syntax

 _expression_. **Creator**

 _expression_Required. A variable that represents a  ** [CustomLabels](407e75b5-4116-fdc7-f0c1-dfd3809cdb41.md)** collection.


## Remarks

If the object was created in Microsoft Word, the  **Creator** property returns the hexadecimal number 4D535744, which represents the string "MSWD." This property was primarily designed to be used on the Macintosh, where each application has a four-character creator code. For example, Microsoft Word has the creator code MSWD. For additional information about this property, consult the language reference Help included with Microsoft Office Macintosh Edition.


## See also


#### Concepts


 [CustomLabels Collection Object](407e75b5-4116-fdc7-f0c1-dfd3809cdb41.md)
#### Other resources


 [CustomLabels Object Members](ee79f452-698d-3089-ed57-b2ca3b125e3d.md)
