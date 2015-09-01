
# OMathBreaks.Creator Property (Word)

 **Last modified:** July 28, 2015

Returns a 32-bit integer that indicates the application in which the add-in was created. Read-only  **Long**.

## Syntax

 _expression_. **Creator**

 _expression_An expression that returns an  ** [OMathBreaks](fa01cd62-b8ad-52bf-f36a-f5d1548d3d1e.md)** object.


## Remarks

If the object was created in Microsoft Word, the  **Creator** property returns the hexadecimal number 4D535744, which represents the string "MSWD." This property was primarily designed to be used on the Macintosh, where each application has a four-character creator code. For example, Microsoft Word has the creator code MSWD. For additional information about this property, consult the language reference Help included with Microsoft Office Macintosh Edition.


 **Note**  This value can also be represented by the constant  **wdCreatorCode**.


## See also


#### Concepts


 [OMathBreaks Collection](fa01cd62-b8ad-52bf-f36a-f5d1548d3d1e.md)
#### Other resources


 [OMathBreaks Object Members](8a16ddcf-9fdc-0cb6-b033-99fe89846a04.md)
