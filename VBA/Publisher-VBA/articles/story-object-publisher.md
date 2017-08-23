---
title: "Объект сценариев (издатель)"
keywords: vbapb10.chm5898239
f1_keywords: vbapb10.chm5898239
ms.prod: publisher
api_name: Publisher.Story
ms.assetid: 0385b4be-0046-9198-a186-0d992601780e
ms.date: 06/08/2017
ms.openlocfilehash: 6bcfdb08e1f554de3d60eea6435f4d0b168fd6d7
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="story-object-publisher"></a>Объект сценариев (издатель)

Представляет текст в кадре несвязанными текст, текст, передаваемых между рамок связанного текста или текста в ячейку таблицы. Объект **материал** является членом **TextFrame** и **TextRange** объекты и коллекции **функциональности** .


## <a name="example"></a>Пример

Используйте свойство **сценариев** для возврата объекта **сценариев** в диапазон текста или текста кадров. В этом примере возвращает История в диапазоне выделенный текст и, если фрагмент текста Вставка текста в диапазон текста.


```
Sub AddTextToStory() 
 With Selection.TextRange.Story 
 If .HasTextFrame Then .TextRange _ 
 .InsertAfter NewText:=vbLf &amp; "This is a test." 
 End With 
End Sub
```

Использование **функциональности** (индекс), где индекс — номер статьи, возвращает объект отдельные **статьи** . В этом примере определяется, если первая статья в активной публикации содержит фрагмент текста и, если это так, форматы абзацы в статье с с половиной дюйм первой строки и отступы шести точки перед все абзацы.




```
Sub StoryParagraphFirstLineIndent() 
 With ActiveDocument.Stories(1) 
 If .HasTextFrame Then 
 With .TextFrame.TextRange.ParagraphFormat 
 .FirstLineIndent = InchesToPoints(0.5) 
 .SpaceBefore = 6 
 End With 
 End If 
 End With 
End Sub
```


## <a name="properties"></a>Properties



|**Name**|
|:-----|
|[Приложения](http://msdn.microsoft.com/library/26c38a3a-e30b-1f2d-d535-57bb978bc4f7%28Office.15%29.aspx)|
|[HasTable](http://msdn.microsoft.com/library/bc4912e2-f521-c6b5-b5a6-a49952014966%28Office.15%29.aspx)|
|[HasTextFrame](http://msdn.microsoft.com/library/10c3a002-05ae-1167-784c-d62066de802d%28Office.15%29.aspx)|
|[Родительский раздел](http://msdn.microsoft.com/library/fbcc74f6-a7ba-df22-0b75-a7b365883d89%28Office.15%29.aspx)|
|[В таблице](http://msdn.microsoft.com/library/e9da80d3-ea3c-b47c-d434-498c72955c14%28Office.15%29.aspx)|
|[TextFrame](http://msdn.microsoft.com/library/bb6ce510-068c-27c2-9df0-a709ab46db2e%28Office.15%29.aspx)|
|[TextRange](http://msdn.microsoft.com/library/c948da79-ea67-0c8c-1df3-2b32499ea9b3%28Office.15%29.aspx)|
|[Type](http://msdn.microsoft.com/library/71e6548d-f54a-b4df-d878-d86a85c1332b%28Office.15%29.aspx)|

