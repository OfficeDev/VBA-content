
# Windows.Creator Property (Word)

 **Last modified:** July 28, 2015

Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only  **Long**.

## Syntax

 _expression_. **Creator**

 _expression_Required. A variable that represents a  ** [Windows](377b493b-e73c-0132-869c-3876c3beaef7.md)** collection.


## Remarks

If the object was created in Microsoft Word, the  **Creator** property returns the hexadecimal number 4D535744, which represents the string "MSWD." This property was primarily designed to be used on the Macintosh, where each application has a four-character creator code. For example, Microsoft Word has the creator code MSWD. For additional information about this property, consult the language reference Help included with Microsoft Office Macintosh Edition.


## See also


#### Concepts


 [Windows Collection Object](377b493b-e73c-0132-869c-3876c3beaef7.md)
#### Other resources


 [Windows Object Members](4a0863e6-b72c-fc50-95ac-3e9a0d231626.md)
