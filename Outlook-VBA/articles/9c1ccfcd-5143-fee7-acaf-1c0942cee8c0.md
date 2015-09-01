
# TaskRequestAcceptItem.BeforeAttachmentPreview Event (Outlook)

 **Last modified:** July 28, 2015

Occurs before an attachment associated with an instance of the parent object is previewed.

## Syntax

 _expression_. **BeforeAttachmentPreview**( **_Attachment_**,  **_Cancel_**)

 _expression_A variable that represents a  **TaskRequestAcceptItem** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|Attachment|Required| ** [Attachment](3e11582b-ac90-0948-bc37-506570bb287b.md)**|The  **Attachment** to be previewed.|
|Cancel|Required| **Boolean**|Set to  **True** to cancel the operation; otherwise, set to **False** to allow the **Attachment** to be previewed.|

## Remarks

This event occurs before an attachment is previewed, either from the attachment strip in the Reading Pane of the active explorer or from the active inspector.


## See also


#### Concepts


 [TaskRequestAcceptItem Object](a2905f72-0a67-b07d-7f85-84fe4de17c25.md)
#### Other resources


 [TaskRequestAcceptItem Object Members](fe91c4cc-f505-11d8-0d0a-84fc4d355651.md)
