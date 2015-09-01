
# ChartTitle.Creator Property (Word)

 **Last modified:** July 28, 2015

Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only  **Long**.

## Syntax

 _expression_. **Creator**

 _expression_A variable that represents a  ** [ChartTitle](fc8ca540-0a29-123b-2fdf-b16aaa1f940c.md)** object.


## Remarks

If the object was created in Microsoft Word, the  **Creator** property returns the hexadecimal number 4D535744, which represents the string "MSWD". This property was primarily designed to be used on the Macintosh, where each application has a four-character creator code. For example, Word has the creator code MSWD. For more information about this property, consult the language reference Help included with Microsoft Office for Mac.


## See also


#### Concepts


 [ChartTitle Object](fc8ca540-0a29-123b-2fdf-b16aaa1f940c.md)
#### Other resources


 [ChartTitle Object Members](e85a7f56-06f4-0561-a37b-7444115965fa.md)
