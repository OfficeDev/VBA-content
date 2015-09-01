
# OlTaskResponse Enumeration (Outlook)

 **Last modified:** July 28, 2015

Indicates the response to a task request.


|**Name**|**Value**|**Description**|
|:-----|:-----|:-----|
| **olTaskAccept**|2|Task accepted.|
| **olTaskAssign**|1|Task reassigned.|
| **olTaskDecline**|3|Task declined.|
| **olTaskSimple**|0|Task is a simple task and cannot be accepted, declined, or assigned. This constant is not a valid parameter to the  **TaskItem.Respond** method.|

## Remarks

Used by the  [TaskItem.ResponseState Property (Outlook)](91f1d4a1-f55b-7379-c1a8-c302bac25a6c.md) and as a parameter to the [TaskItem.Respond Method (Outlook)](1befabf7-262f-897a-d1dc-49be4e7ddf9b.md).

