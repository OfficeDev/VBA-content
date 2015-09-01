
# Editors.Creator Property (Word)

 **Last modified:** July 28, 2015

Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only  **Long**.

## Syntax

 _expression_. **Creator**

 _expression_Required. A variable that represents an  ** [Editors](acce718a-e3c1-deac-8b7f-fd8a5a9e47c6.md)** collection.


## Remarks

If the object was created in Microsoft Word, the  **Creator** property returns the hexadecimal number 4D535744, which represents the string "MSWD." This property was primarily designed to be used on the Macintosh, where each application has a four-character creator code. For example, Microsoft Word has the creator code MSWD. For additional information about this property, consult the language reference Help included with Microsoft Office Macintosh Edition.


## See also


#### Concepts


 [Editors Collection](acce718a-e3c1-deac-8b7f-fd8a5a9e47c6.md)
#### Other resources


 [Editors Object Members](dcb26f83-bbff-8d3a-2493-f7d87ce40d21.md)
