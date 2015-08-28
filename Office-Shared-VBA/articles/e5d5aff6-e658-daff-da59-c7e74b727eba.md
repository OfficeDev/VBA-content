
# SharedWorkspaceFiles.Creator Property (Office)

 **Last modified:** July 28, 2015

Gets a 32-bit integer that indicates the application in which the  **SharedWorkspaceFiles** object was created. Read-only.

 **Note**  Beginning with Microsoft Office 2010, this object or member has been deprecated and should not be used.


## Syntax

 _expression_. **Creator**

 _expression_A variable that represents a  **SharedWorkspaceFiles** object.


### Return Value

Long


## Remarks

As an example, if the object was created in Microsoft Word, this property returns 1297307460, which represents the string "MSWD"; in Microsoft Excel, this property returns 1480803660. This value can also be represented by the constant wdCreatorCode in Word, or xlCreatorCode in Excel. The  **Creator** property was primarily designed to be used on the Macintosh, where each application has a four-character creator code. For example, Microsoft Word has the creator code MSWD. For additional information about this property, consult the language reference Help included with Microsoft Office Macintosh Edition.

The  **Creator** property always returns the numeric identifier for the active application, just as the **Application** property always returns the name of the active applicatin in string form. Use the **CreatedBy** property of the **SharedWorkspaceFile** object to return the name of the individual who created the object. Use document properties to return information about the authors of Office documents.


## See also


#### Concepts


 [SharedWorkspaceFiles Object](5e2937f7-f794-dffb-a1ec-69ea9a9e3546.md)
#### Other resources


 [SharedWorkspaceFiles Object Members](30e841ce-c8f1-249a-3bc7-6f204be64536.md)
