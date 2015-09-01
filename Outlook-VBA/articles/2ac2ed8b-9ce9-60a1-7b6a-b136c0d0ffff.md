
# CardView.Filter Property (Outlook)

 **Last modified:** July 28, 2015

 **In this article**
 [Syntax](#sectionSection0)
 [Remarks](#sectionSection1)
 [Example](#sectionSection2)


Returns or sets a  **String** value that represents the filter for a view. Read/write.


## Syntax
<a name="sectionSection0"> </a>

 _expression_. **Filter**

 _expression_A variable that represents a  **CardView** object.


## Remarks
<a name="sectionSection1"> </a>

The value of this property is a string, in DAV Searching and Locating (DASL) syntax, that represents the current filter for the view. For more information about using DASL syntax to filter items in a view, see  [Filtering Items](4038e042-1b07-5d18-18b0-c2b58c9c42da.md).


## Example
<a name="sectionSection2"> </a>

The following Visual Basic for Applications (VBA) example obtains a  ** [View](41c8d149-9912-1685-4c8b-3c849cc6f1ed.md)** object by using the ** [CurrentView](177e6387-9ccb-cb71-bbe5-332c25485848.md)** property of the ** [Explorer](026591e5-049f-503a-4166-34e6dbc225fb.md)** object, then sets the ** [Filter](9a4b4b27-d543-df82-3058-e0a6ad2f51a1.md)** property of the **View** object to display only those Outlook items that were received last week.


```
Private Sub FilterViewToLastWeek() 
 
 Dim objView As View 
 
 
 
 ' Obtain a View object reference to the current view. 
 
 Set objView = Application.ActiveExplorer.CurrentView 
 
 
 
 ' Set a DASL filter string, using a DASL macro, to show 
 
 ' only those items that were received last week. 
 
 objView.Filter = "%lastweek(""urn:schemas:httpmail:datereceived"")%" 
 
 
 
 ' Save and apply the view. 
 
 objView.Save 
 
 objView.Apply 
 
End Sub 
 

```


## See also
<a name="sectionSection2"> </a>


#### Concepts


 [CardView Object](cdac229b-f2b6-9ecb-e1a7-b53509426570.md)
#### Other resources


 [CardView Object Members](8b9eda10-1ece-c961-e432-3fca6dfb4f07.md)
