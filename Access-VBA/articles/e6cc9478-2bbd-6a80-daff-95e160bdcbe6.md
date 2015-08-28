
# Report.FontItalic Property (Access)

 **Last modified:** July 28, 2015

You can use the  **FontItalic** property to specify whether text is italic in the following situations:

- When displaying or printing controls on forms and reports.
    
- When using the  **Print**method on a report.
    
 Read/write **Boolean**.

## Syntax

 _expression_. **FontItalic**

 _expression_A variable that represents a  **Report** object.


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


 [Report Object](6f77c1b4-a9ce-7caa-204c-fe0755c6f9df.md)
#### Other resources


 [Report Object Members](73370a33-1ca0-da4d-9e36-88011bc2b93e.md)
