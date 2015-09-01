
# MailItem.GetInspector Property (Outlook)

 **Last modified:** July 28, 2015

 **In this article**
 [Syntax](#sectionSection0)
 [Remarks](#sectionSection1)
 [Example](#sectionSection2)


Returns an  ** [Inspector](d7384756-669c-0549-1032-c3b864187994.md)** object that represents an inspector initialized to contain the specified item. Read-only.


## Syntax
<a name="sectionSection0"> </a>

 _expression_. **GetInspector**

 _expression_A variable that represents a  **MailItem** object.


## Remarks
<a name="sectionSection1"> </a>

This property is useful for returning an  **Inspector** object in which to display the item, as opposed to using the ** [Application.ActiveInspector](3f2b6491-7b4b-8165-327e-b319711d5656.md)** method and setting the ** [Inspector.CurrentItem](eaaf0192-a169-c107-95a6-b8e759a3b873.md)** property. If an **Inspector** object already exists for the item, the **GetInspector** property will return that **Inspector** object instead of creating a new one.


## Example
<a name="sectionSection2"> </a>

This Visual Basic for Applications (VBA) example shows a function  `InsertBodyTextInWordEditor` that creates a mail item, assigns it a title and adds text for the body. The function sets the ** [Subject](5f3e465d-ac2b-a573-0e85-1134e65df017.md)** property to assign the title "Testing...". It then calls the ** [Display](19ead642-b7bd-579f-e43b-ef5c5d0cfecb.md)** method to open the mail item in an inspector. To insert text in a Word editor as the body of the mail item, the function uses the ** [Document](http://msdn.microsoft.com/library/8d83487a-2345-a036-a916-971c9db5b7fb%28Office.15%29.aspx)** object and ** [Range](http://msdn.microsoft.com/library/15a7a1c4-5f3f-5b6e-60e9-29688de3f274%28Office.15%29.aspx)** object in the Word object model. The function uses the item's **GetInspector** property to get the existing **Inspector** object, and then uses the ** [Inspector.WordEditor](9e09b772-f679-19e6-905e-552ccadb0d24.md)** property to obtain a **Word.Document** object for the item. Using the **Word.Document** object, the function accesses the **Word.Range** object and inserts text into the body of the item.

Since this example accesses the Word object model, you must first add a reference to the Microsoft Word Object Library to compile the example successfully.




```
Sub InsertBodyTextInWordEditor() 
 Dim myItem As Outlook.MailItem 
 Dim myInspector As Outlook.Inspector 
 'You must add a reference to the Microsoft Word Object Library 
 'before this sample will compile 
 Dim wdDoc As Word.Document 
 Dim wdRange As Word.Range 
 
 On Error Resume Next 
 Set myItem = Application.CreateItem(olMailItem) 
 myItem.Subject = "Testing..." 
 myItem.Display 
 'GetInspector property returns Inspector 
 Set myInspector = myItem.GetInspector 
 'Obtain the Word.Document for the Inspector 
 Set wdDoc = myInspector.WordEditor 
 If Not (wdDoc Is Nothing) Then 
 'Use the Range object to insert text 
 Set wdRange = wdDoc.Range(0, wdDoc.Characters.Count) 
 wdRange.InsertAfter ("Hello world!") 
 End If 
End Sub
```


## See also
<a name="sectionSection2"> </a>


#### Concepts


 [MailItem Object](14197346-05d2-0250-fa4c-4a6b07daf25f.md)
#### Other resources


 [MailItem Object Members](1094d7df-ee80-a4b0-5a21-db2979506e6b.md)
