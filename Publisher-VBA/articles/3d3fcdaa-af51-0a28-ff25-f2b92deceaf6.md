
# Application.MailMergeWizardStateChange Event (Publisher)

 **Last modified:** July 28, 2015

 **In this article**
 [Syntax](#sectionSection0)
 [Remarks](#sectionSection1)
 [Example](#sectionSection2)


Occurs when a user changes from a specified step to a specified step in the Mail Merge Wizard.


## Syntax
<a name="sectionSection0"> </a>

 _expression_. **MailMergeWizardStateChange**( **_Doc_**,  **_FromState_**)

 _expression_A variable that represents an  **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|Doc|Required| **Document**|The mail merge main document.|
|FromState|Required| **Integer**|The Mail Merge Wizard step from which a user is moving.|

## Remarks
<a name="sectionSection1"> </a>

To access the  **Application** object events, declare an **Application** object variable in the General Declarations section of a code module. Then set the variable equal to the **Application** object for which you want to access events.


## Example
<a name="sectionSection2"> </a>

This example displays a message when a users moves from step three of the Mail Merge Wizard to step four. Based on the user's answer to the message, the user will either continue on to step four or return to step three.


```
Private Sub MailMergeApp_MailMergeWizardStateChange(ByVal Doc As Document, _ 
 ByVal FromState As Long) 
 
 Select Case FromState 
 Case 1 
 MsgBox "Now you will build your publication merge " &amp; _ 
 "by adding fields to your publication." 
 Case 2 
 MsgBox "Now you will see your publication " &amp; _ 
 "merged with the records in the data source." 
 Case 3 
 MsgBox "Now you will complete the mail merge process." 
 End Select 
 
End Sub
```


## See also
<a name="sectionSection2"> </a>


#### Concepts


 [Application Object](acfc7efb-e6a5-a89a-3aee-3cb4af2f3508.md)
#### Other resources


 [Application Object Members](aa4d515b-f779-b8b5-968a-8e5f7466fb56.md)
