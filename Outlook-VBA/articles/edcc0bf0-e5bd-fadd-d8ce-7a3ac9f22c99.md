
# TaskRequestItem.BeforeAttachmentWriteToTempFile Event (Outlook)

 **Last modified:** July 28, 2015

Occurs before an attachment associated with an instance of the parent object is written to a temporary file.

## Syntax

 _expression_. **BeforeAttachmentWriteToTempFile**( **_Attachment_**,  **_Cancel_**)

 _expression_A variable that represents a  **TaskRequestItem** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|Attachment|Required| ** [Attachment](3e11582b-ac90-0948-bc37-506570bb287b.md)**|The  **Attachment** to be written.|
|Cancel|Required| **Boolean**|Set to  **True** to cancel the operation; otherwise, set to **False** to allow the **Attachment** to be written.|

## See also


#### Concepts


 [TaskRequestItem Object](2908a28a-634c-e786-aa53-f3e32038b727.md)
#### Other resources


 [TaskRequestItem Object Members](d43114ee-be91-ff02-3424-525da2cf3a50.md)
