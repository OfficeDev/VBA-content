
# OMathRad.Creator Property (Word)

 **Last modified:** July 28, 2015

Returns a 32-bit integer that indicates the application in which the add-in was created. Read-only  **Long**.

## Syntax

 _expression_. **Creator**

 _expression_An expression that returns an  ** [OMathRad](2179cda9-b1dc-9593-c4f9-99496081e191.md)** object.


## Remarks

If the object was created in Microsoft Word, the  **Creator** property returns the hexadecimal number 4D535744, which represents the string "MSWD." This property was primarily designed to be used on the Macintosh, where each application has a four-character creator code. For example, Microsoft Word has the creator code MSWD. For additional information about this property, consult the language reference Help included with Microsoft Office Macintosh Edition.


 **Note**  This value can also be represented by the constant  **wdCreatorCode**.


## See also


#### Concepts


 [OMathRad Object](2179cda9-b1dc-9593-c4f9-99496081e191.md)
#### Other resources


 [OMathRad Object Members](90b7d5a3-f695-9ae3-0cde-c86fbd94352f.md)
