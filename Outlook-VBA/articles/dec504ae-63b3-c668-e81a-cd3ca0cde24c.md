
# TaskItem.BeforeAttachmentAdd Event (Outlook)

 **Last modified:** July 28, 2015

Occurs before an attachment is added to an instance of the parent object.

## Syntax

 _expression_. **BeforeAttachmentAdd**( **_Attachment_**,  **_Cancel_**)

 _expression_A variable that represents a  **TaskItem** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|Attachment|Required| ** [Attachment](3e11582b-ac90-0948-bc37-506570bb287b.md)**|The  **Attachment** to be added to the item.|
|Cancel|Required| **Boolean**|Set to  **True** to cancel the operation; otherwise, set to **False** to allow the **Attachment** to be added.|

## See also


#### Concepts


 [TaskItem Object](5df8cfa5-5460-a5a1-a130-ba5bca1a0091.md)
#### Other resources


 [TaskItem Object Members](97234a76-2fc5-bbe4-2e14-25ae18694fc9.md)
