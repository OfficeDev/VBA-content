---
title: Creating Frames Pages
keywords: vbawd10.chm5285976
f1_keywords:
- vbawd10.chm5285976
ms.prod: word
ms.assetid: 0245564e-b2df-83cd-1e32-e63079970dc1
ms.date: 06/08/2017
---


# Creating Frames Pages

In Word, you can use frames in your Web page design to make your information organized and easy to access. A frames page, also called a frameset, is a Web page that is divided into two or more sections, each of which points to another Web page. A frame on a frames page can also point to another frames page. For information about creating frames and frames pages in the Word user interface, see Word Help.

Frames and frames pages are created with a series of HTML tags. The Visual Basic object model for working with frames and frames pages is best understood by examining their HTML tags.

## Frames pages in HTML

In HTML, frames pages and the frames they contain are built using a hierarchical set of <FRAMESET> and <FRAME> tags. A frameset can contain both frames and other framesets. For example, the following HTML creates a frameset with a frame on top and a frameset immediately below it. That frameset contains a frame on the left and a frameset on the right. That frameset contains two frames, one on top of the other.


```XML
<FRAMESET ROWS="100, *"> 
    <FRAME NAME=top SRC="banner.htm"> 
    <FRAMESET COLS="20%, *"> 
        <FRAME NAME=left SRC="contents.htm"> 
        <FRAMESET ROWS="75%, *"> 
            <FRAME NAME=main SRC="main.htm"> 
            <FRAME NAME=bottom SRC="footer.htm"> 
        </FRAMESET> 
    </FRAMESET> 
</FRAMESET>
```


 **Note**  To better understand the preceding HTML example, paste the example into a blank text document, rename the document "framespage.htm", and open the document in Word or in a Web browser.


## The Frameset Object

The  **[Frameset](frameset-object-word.md)** object encompasses the functionality of both tags. Each **Frameset** object is either of type **wdFramesetTypeFrameset** or **wdFramesetTypeFrame**, which represent the HTML tags <FRAMESET> and <FRAME> respectively. Properties beginning with "Frameset" apply to  **Frameset** objects of type **wdFramesetTypeFrameset** ( **[FramesetBorderColor](frameset-framesetbordercolor-property-word.md)** and **[FramesetBorderWidth](frameset-framesetborderwidth-property-word.md)** . Properties beginning with "Frame" apply to **Frameset** objects of type **wdFramesetTypeFrame** ( **[FrameDefaultURL](frameset-framedefaulturl-property-word.md)**,  **[FrameDisplayBorders](frameset-framedisplayborders-property-word.md)**,  **[FrameLinkToFile](frameset-framelinktofile-property-word.md)**,  **[FrameName](frameset-framename-property-word.md)**,  **[FrameResizable](frameset-frameresizable-property-word.md)**, and  **[FrameScrollBarType](frameset-framescrollbartype-property-word.md)**).


## Traversing the Frameset Object Hierarchy

Because frames pages are defined as a hierarchical set of HTML tags, the object model for accessing  **Frameset** objects is also hierarchical. Use the **[ChildFramesetItem](frameset-childframesetitem-property-word.md)** and **[ParentFrameset](frameset-parentframeset-property-word.md)** properties to traverse the hierarchy of **Frameset** objects. For example,


```
MyFrameset.ChildFramesetItem(n)
```

returns a  **Frameset** object corresponding to the _n_th first-level <FRAMESET> or <FRAME> tag between the <FRAMESET> and </FRAMESET> tags corresponding to  `MyFrameset`.

If  `MyFrameset` is a **Frameset** object corresponding to the outermost <FRAMESET> tags in the preceding HTML example, `MyFrameset.ChildFramesetItem(1)` returns a **Frameset** object of type **wdFramesetTypeFrame** that corresponds to the frame named "top." Similarly, `MyFrameset.ChildFramesetItem(2)` returns a **Frameset** object of type **wdFramesetTypeFrameset**, itself containing two  **Frameset** objects: the first object corresponds to the frame named "left," the second is of type **wdFramesetTypeFrameset**.

 **Frameset** objects of type **wdFramesetTypeFrame** have no child frames, while objects of **wdFramesetTypeFrameset** have at least one.

The following Visual Basic example displays the names of the four frames defined in the preceding HTML example.




```vb
Dim MyFrameset As Frameset 
Dim Name1 As String 
Dim Name2 As String 
Dim Name3 As String 
Dim Name4 As String 
 
Set MyFrameset = ActiveWindow.Document.Frameset 
 
With MyFrameset 
    Name1 = .ChildFramesetItem(1).FrameName 
    With .ChildFramesetItem(2) 
        Name2 = .ChildFramesetItem(1).FrameName 
        With .ChildFramesetItem(2) 
            Name3 = .ChildFramesetItem(1).FrameName 
            Name4 = .ChildFramesetItem(2).FrameName 
        End With 
    End With 
End With 
 
Debug.Print Name1, Name2, Name3, Name4
```


## Individual Frames and the Entire Frames Page

To return the  **Frameset** object associated with a particular frame on a frames page, use the **[Frameset](pane-frameset-property-word.md)** property of a **[Pane](pane-object-word.md)** object. For example,


```vb
ActiveWindow.Panes(1).Frameset
```

returns the  **Frameset** object that corresponds to the first frame of the frames page.

The frames page is itself a document separate from the documents that make up the content of the individual frames. The  **Frameset** object associated with a frames page is accessed from its corresponding **[Document](document-object-word.md)** object, which in turn is accessed from the **[Window](window-object-word.md)** object in which the frames page appears. For example,




```vb
ActiveWindow.Document.Frameset
```

returns the  **Frameset** object for the frames page in the active window.


 **Note**  When working with frames pages, the  **[ActiveDocument](application-activedocument-property-word.md)** property returns the **Document** object associated with the frame in the active pane, not the entire frames page.


## Creating a Frames Page and Its Content from Scratch

This example creates a new frames page with three frames, adds text to each frame, and sets the background color for each frame. It inserts two hyperlinks into the Left frame: the first hyperlink opens a document called One.htm in the Main frame, and the second opens a document called Two.htm in the entire window. For these hyperlinks to work, you must create files called One.htm and Two.htm or change the strings to the names of existing files.


 **Note**  As each frame is created, Word creates a new document whose content will be loaded into the new frame. The example saves the frames page which automatically saves the documents associated with each of the three frames.


```vb
Sub FramesetExample1() 
 
    ' Create new frames page. 
    Documents.Add DocumentType:=wdNewFrameset 
 
    With ActiveWindow 
        ' Add text and color to first frame. 
        Selection.TypeText Text:="BANNER FRAME" 
        With ActiveDocument.Background.Fill 
            .ForeColor.RGB = RGB(204, 153, 255) 
            .Visible = msoTrue 
        End With 
 
        ' Add new frame below top frame. 
        .ActivePane.Frameset.AddNewFrame _ 
            wdFramesetNewFrameBelow 
        ' Add text and color to bottom frame. 
        .ActivePane.Frameset.FrameName = "main" 
        Selection.TypeText Text:="MAIN FRAME" 
        With ActiveDocument.Background.Fill 
            .ForeColor.RGB = RGB(0, 128, 128) 
            .Visible = msoTrue 
        End With 
 
        ' Add new frame to left of bottom frame. 
        .ActivePane.Frameset.AddNewFrame _ 
            wdFramesetNewFrameLeft 
        ' Set the width to 25% of the window width. 
        With .ActivePane.Frameset 
            .WidthType = wdFramesetSizeTypePercent 
            .Width = 25 
            .FrameName = "left" 
        End With 
        ' Add text and color to left frame. 
        Selection.TypeText Text:="LEFT FRAME" 
        With ActiveDocument.Background.Fill 
            .ForeColor.RGB = RGB(204, 255, 255) 
            .Visible = msoTrue 
        End With 
        Selection.TypeParagraph 
        Selection.TypeParagraph 
        ' Add hyperlinks to left frame. 
        With ActiveDocument.Hyperlinks 
            .Add Anchor:=Selection.Range, _ 
                Address:="one.htm", Target:="main" 
            Selection.TypeParagraph 
            Selection.TypeParagraph 
            .Add Anchor:=Selection.Range, _ 
                Address:="two.htm", Target:="_top" 
        End With 
        
        ' Activate top frame. 
        .Panes(1).Activate 
        ' Set the height to 1 inch. 
        With .ActivePane.Frameset 
            .HeightType = wdFramesetSizeTypeFixed 
            .Height = InchesToPoints(1) 
            .FrameName = "top" 
        End With 
 
        ' Save the frames page and its associated files. 
        .Document.SaveAs FileName:="default.htm", _ 
            FileFormat:=wdFormatHTML 
    End With 
 
End Sub
```


## Creating a Frames Page that Displays Content from Existing Files

This example creates a frames page similar to the one above, but sets the default URL for each frame to an existing document so that the content of that document is displayed in the frame. For this example to work, you must create files called Main.htm, Left.htm, and Banner.htm or change the strings in the example to the names of existing files.


```vb
Sub FramesetExample2() 
    
    ' Create new frames page. 
    Documents.Add DocumentType:=wdNewFrameset 
 
    With ActiveWindow 
        ' Add new frame below top frame. 
        .ActivePane.Frameset.AddNewFrame _ 
            wdFramesetNewFrameBelow 
        ' Set the name and initial page for the frame. 
        With .ActivePane.Frameset 
            .FrameName = "main" 
            .FrameDefaultURL = "main.htm" 
        End With 
        
        ' Add new frame to left of bottom frame. 
        .ActivePane.Frameset.AddNewFrame _ 
            wdFramesetNewFrameLeft 
        With .ActivePane.Frameset 
            ' Set the width to 25% of the window width. 
            .WidthType = wdFramesetSizeTypePercent 
            .Width = 25 
            ' Set the name and initial page for the frame. 
            .FrameName = "left" 
            .FrameDefaultURL = "left.htm" 
        End With 
    
        ' Activate top frame. 
        .Panes(1).Activate 
        With .ActivePane.Frameset 
            ' Set the height to 1 inch. 
            .HeightType = wdFramesetSizeTypeFixed 
            .Height = InchesToPoints(1) 
            ' Set the name and initial page for the frame. 
            .FrameName = "top" 
            .FrameDefaultURL = "banner.htm" 
        End With 
 
        ' Save the frameset. 
        .Document.SaveAs FileName:="default.htm", _ 
            FileFormat:=wdFormatHTML 
    End With 
 
End Sub
```


