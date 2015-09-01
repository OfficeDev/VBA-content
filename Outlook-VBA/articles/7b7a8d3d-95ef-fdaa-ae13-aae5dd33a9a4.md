
# PostItem.ReadComplete Event (Outlook)
Occurs when Outlook has completed reading the properties of the item.

 **Last modified:** July 28, 2015

 **In this article**
 [Version information](#sectionSection0)
 [Syntax](#sectionSection1)
 [Remarks](#sectionSection2)


## Version information
<a name="sectionSection0"> </a>

Version Added: Outlook 2013 


## Syntax
<a name="sectionSection1"> </a>

 _expression_. **ReadComplete**(Cancel)

 _expression_A variable that represents a  **PostItem** object.


### Parameters



|**Name**|**Required/Optional**|**Data type**|**Description**|
|:-----|:-----|:-----|:-----|
|||||
|Cancel|Required| **Boolean**|(Not used in VBScript).  **False** when the event occurs. If the event procedure sets this argument to **True**, the read operation is not completed and the item is not displayed in the Reading Pane or inspector.|

## Remarks
<a name="sectionSection2"> </a>

The  **ReadComplete** event occurs after the [BeforeRead](26a64e4e-a48e-84e8-4fea-70913a8f170f.md) event and before the [Read](404c9b17-c5b6-a802-420a-f8fd279b5f9b.md) event for the item.

To determine when the item is unloaded from memory, use the  [Unload](42dea931-c3dd-b8ff-5ace-0744b17650e0.md) event.

The  **ReadComplete** event corresponds to the Exchange Client Extensions (ECE) event **IExchExtMessageEvents::OnReadComplete**.


## See also
<a name="sectionSection2"> </a>


#### Concepts


 [PostItem Object](de44065d-4e93-315a-279f-7b92f09c0465.md)
#### Other resources


 [PostItem Object Members](5b150db1-c96d-0721-ec36-d5b5ebc20fd8.md)
