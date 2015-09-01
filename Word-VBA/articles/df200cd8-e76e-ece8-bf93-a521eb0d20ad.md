
# UpBars.Creator Property (Word)

 **Last modified:** July 28, 2015

Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only  **Long**.

## Syntax

 _expression_. **Creator**

 _expression_A variable that represents an  ** [UpBars](22dff1d2-8f1b-8c48-354c-570906e0f830.md)** object.


## Remarks

If the object was created in Microsoft Word, the  **Creator** property returns the hexadecimal number 4D535744, which represents the string "MSWD". This property was primarily designed to be used on the Macintosh, where each application has a four-character creator code. For example, Word has the creator code MSWD. For more information about this property, consult the language reference Help included with Microsoft Office for Mac.


## See also


#### Concepts


 [UpBars Object](22dff1d2-8f1b-8c48-354c-570906e0f830.md)
#### Other resources


 [UpBars Object Members](7772742e-1230-6987-f8f3-f3663ea4329b.md)
