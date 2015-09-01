
# OMathLimUpp.Creator Property (Word)

 **Last modified:** July 28, 2015

Returns a 32-bit integer that indicates the application in which the add-in was created. Read-only  **Long**.

## Syntax

 _expression_. **Creator**

 _expression_An expression that returns an  ** [OMathLimUpp](3c7ca001-8533-52c9-5343-8a89892c0a16.md)** object.


## Remarks

If the object was created in Microsoft Word, the  **Creator** property returns the hexadecimal number 4D535744, which represents the string "MSWD." This property was primarily designed to be used on the Macintosh, where each application has a four-character creator code. For example, Microsoft Word has the creator code MSWD. For additional information about this property, consult the language reference Help included with Microsoft Office Macintosh Edition.


 **Note**  This value can also be represented by the constant  **wdCreatorCode**.


## See also


#### Concepts


 [OMathLimUpp Object](3c7ca001-8533-52c9-5343-8a89892c0a16.md)
#### Other resources


 [OMathLimUpp Object Members](789004f4-1c6e-de7e-484b-7da6a9d185fd.md)
