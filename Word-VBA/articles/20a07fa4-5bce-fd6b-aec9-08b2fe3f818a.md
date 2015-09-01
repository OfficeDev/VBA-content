
# HeadersFooters.Creator Property (Word)

 **Last modified:** July 28, 2015

Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only  **Long**.

## Syntax

 _expression_. **Creator**

 _expression_Required. A variable that represents a  ** [HeadersFooters](41dbbaa7-f139-3d3c-54d4-03a57ab8417a.md)** collection.


## Remarks

If the object was created in Microsoft Word, the  **Creator** property returns the hexadecimal number 4D535744, which represents the string "MSWD." This property was primarily designed to be used on the Macintosh, where each application has a four-character creator code. For example, Microsoft Word has the creator code MSWD. For additional information about this property, consult the language reference Help included with Microsoft Office Macintosh Edition.


## See also


#### Concepts


 [HeadersFooters Collection Object](41dbbaa7-f139-3d3c-54d4-03a57ab8417a.md)
#### Other resources


 [HeadersFooters Object Members](6cf7f768-d356-7ff4-089c-7f3f810d00a8.md)
