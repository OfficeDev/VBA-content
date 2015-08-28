
# FormatCondition.FontItalic Property (Access)

 **Last modified:** July 28, 2015

You can use the  **FontItalic** property to specify whether text is italic in the following situations:

- When displaying or printing controls on forms and reports.
    
- When using the  **Print**method on a report.
    
 Read/write **Boolean**.

## Syntax

 _expression_. **FontItalic**

 _expression_A variable that represents a  **FormatCondition** object.


## Remarks

The  **FontItalic** property uses the following settings.



|**Setting**|**Description**|
|:-----|:-----|
| **True**|The text is italic.|
| **False**|(Default) The text isn't italic.|
For reports, you can use this property only in an event procedure or in a macro specified by the  **OnPrint**event property setting.

You can set the default for this property by using the default control style or the  **DefaultControl**property in Visual Basic.


## See also


#### Concepts


 [FormatCondition Object](a31deaae-b32d-c45b-b3b2-113a9e62cc7a.md)
#### Other resources


 [FormatCondition Object Members](98a01bf0-3d5c-5ea4-9291-97ddd24fd7a1.md)
