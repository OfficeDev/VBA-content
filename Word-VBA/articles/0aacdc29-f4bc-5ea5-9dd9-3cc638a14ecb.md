
# DownBars.Creator Property (Word)

 **Last modified:** July 28, 2015

Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only  **Long**.

## Syntax

 _expression_. **Creator**

 _expression_A variable that represents a  ** [DownBars](d0cf170e-0c58-2d01-a4b2-1eaf65dbfa3c.md)** object.


## Remarks

If the object was created in Microsoft Word, the  **Creator** property returns the hexadecimal number 4D535744, which represents the string "MSWD". This property was primarily designed to be used on the Macintosh, where each application has a four-character creator code. For example, Word has the creator code MSWD. For more information about this property, consult the language reference Help included with Microsoft Office for Mac.


## See also


#### Concepts


 [DownBars Object](d0cf170e-0c58-2d01-a4b2-1eaf65dbfa3c.md)
#### Other resources


 [DownBars Object Members](ed402462-03fc-d980-6f8d-b3e7ae72aee2.md)
