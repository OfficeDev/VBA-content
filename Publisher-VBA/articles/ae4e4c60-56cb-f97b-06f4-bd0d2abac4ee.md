
# Attachments.ClearAll Method (Publisher)

 **Last modified:** July 28, 2015

 **In this article**
 [Syntax](#sectionSection0)
 [Remarks](#sectionSection1)
 [Example](#sectionSection2)


Clears (deletes) all the  **Attachment** objects in the parent **Attachments** collection of an e-mail merge message.


## Syntax
<a name="sectionSection0"> </a>

 _expression_. **ClearAll**

 _expression_A variable that represents an  **Attachments** collection.


## Remarks
<a name="sectionSection1"> </a>

To clear an individual attachment, use the  ** [Delete](935fa9e7-9d40-b820-e386-1a1960845da1.md)** method of the specific **Attachment** object


## Example
<a name="sectionSection2"> </a>

The following Microsoft Visual Basic for Applications (VBA) macro shows how to clear all the attachment to the message in an e-mail merge. The code prints the number of current attachments to the message in the  **Immediate** window and then deletes all of the **Attachment** objects in the collection.


```
Public Sub ClearAll_Example() 
 
 Dim pubAttachments As Publisher.Attachments 
 
 Dim pubMailMerge As Publisher.MailMerge 
 Dim pubEmailMergeEnvelope As Publisher.EmailMergeEnvelope 
 
 Set pubMailMerge = ThisDocument.MailMerge 
 Set pubEmailMergeEnvelope = pubMailMerge.EmailMergeEnvelope 
 Set pubAttachments = pubEmailMergeEnvelope.Attachemts 
 
 Debug.Print pubAttachments.Count 
 pubAttachments.ClearAll 
 
End Sub
```


## See also
<a name="sectionSection2"> </a>


#### Concepts


 [Attachments Collection](61957961-8c75-992f-159c-51412ed309ea.md)
#### Other resources


 [Attachments Object Members](fbc479ab-ac16-7ee6-f585-5fe63f66b757.md)
