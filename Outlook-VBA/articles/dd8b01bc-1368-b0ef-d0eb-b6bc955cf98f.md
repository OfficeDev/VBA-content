
# TaskRequestDeclineItem.BeforeCheckNames Event (Outlook)

 **Last modified:** July 28, 2015

Occurs just before Microsoft Outlook starts resolving names in the recipient collection for an item (which is an instance of the parent object).

## Syntax

 _expression_. **BeforeCheckNames**( **_Cancel_**)

 _expression_A variable that represents a  **TaskRequestDeclineItem** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|Cancel|Required| **Boolean**| **False** when the event occurs. If the event procedure sets this argument to **True**, the name resolution process is not completed.|

## Remarks

You use the  **BeforeCheckNames** event in VBScript, but the event does not fire when an e-mail name is resolved on the form.

The event does not fire under the following circumstances:


- You customized a Journal Entry form and then resolved a contact in the  **Contacts** field.
    
- You customized a Contact form and then resolved a contact in the  **Contacts** field.
    
- You customized any type of form and Outlook automatically resolved the name in the background.
    
- You programmatically created and resolved a recipient.
    



## See also


#### Concepts


 [TaskRequestDeclineItem Object](e842c7c0-7943-9219-329b-30b892ab99b0.md)
#### Other resources


 [TaskRequestDeclineItem Object Members](3de31d0d-2444-876c-5d4d-1192851301af.md)
