
# FileConverter.CanOpen Property (Word)

 **Last modified:** July 28, 2015

 **In this article**
 [Syntax](#sectionSection0)
 [Remarks](#sectionSection1)
 [Example](#sectionSection2)


 **True** if the specified file converter is designed to open files. Read-only **Boolean**.


## Syntax
<a name="sectionSection0"> </a>

 _expression_. **CanOpen**

 _expression_A variable that represents a  ** [FileConverter](41af2a9b-75cc-253d-4954-4fb42c88530f.md)** object.


## Remarks
<a name="sectionSection1"> </a>

The  ** [CanSave](a1de7523-5b9c-b606-4308-9445e3c4c76d.md)**property returns  **True** if the specified file converter can be used to save (export) files.


## Example
<a name="sectionSection2"> </a>

This example determines whether the first file converter is able to open files.


```
If FileConverters(1).CanOpen = True Then 
 MsgBox FileConverters(1).FormatName &amp; " can open files" 
End If
```

This example determines whether the WordPerfect6x file converter can be used to open files. If the CanOpen property returns True, a document named "Test.wp" is opened.




```
If FileConverters("WordPerfect6x").CanOpen = True Then 
 Documents.Open FileName:="C:\Test.wp", _ 
 Format:=FileConverters("WordPerfect6x").OpenFormat 
End If
```


## See also
<a name="sectionSection2"> </a>


#### Concepts


 [FileConverter Object](41af2a9b-75cc-253d-4954-4fb42c88530f.md)
#### Other resources


 [FileConverter Object Members](cdf7a124-6c27-0edf-7a29-1b28f70d834f.md)
