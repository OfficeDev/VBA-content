---
title: "Объект гиперссылки (издатель)"
keywords: vbapb10.chm6946815
f1_keywords: vbapb10.chm6946815
ms.prod: publisher
api_name: Publisher.Hyperlinks
ms.assetid: a82724b9-e792-b0e6-d1c3-25ce6021ad29
ms.date: 06/08/2017
ms.openlocfilehash: 69e01a384fc7b4ebc25814df0443e84b91f693aa
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="hyperlinks-object-publisher"></a>Объект гиперссылки (издатель)

Представляет коллекцию объектов **[гиперссылок](hyperlink-object-publisher.md)** в диапазон текста.


## <a name="example"></a>Пример

Свойство **[гиперссылки](http://msdn.microsoft.com/library/0cf1f043-532c-3ffc-67cf-389adc5ac02f%28Office.15%29.aspx)** используется для возврата коллекции **гиперссылки** . В следующем примере удаляются все текст гиперссылки в активной публикации, содержащие слово «Tailspin» в поле адрес.


```
Sub DeleteMSHyperlinks() 
 Dim pgsPage As Page 
 Dim shpShape As Shape 
 Dim hprLink As Hyperlink 
 For Each pgsPage In ActiveDocument.Pages 
 For Each shpShape In pgsPage.Shapes 
 If shpShape.HasTextFrame = msoTrue Then 
 If shpShape.TextFrame.HasText = msoTrue Then 
 For Each hprLink In shpShape.TextFrame.TextRange.Hyperlinks 
 If InStr(hprLink.Address, "tailspin") <> 0 Then 
 hprLink.Delete 
 Exit For 
 End If 
 Next 
 Else 
 shpShape.Hyperlink.Delete 
 End If 
 End If 
 Next 
 Next 
End Sub
```

Используйте метод **[Add](http://msdn.microsoft.com/library/f5a8cc01-a571-623d-bfab-fe48e43a21b1%28Office.15%29.aspx)** для создания гиперссылки и добавить его в коллекцию **гиперссылки** . В следующем примере создается новый гиперссылки для указанного веб-сайта.




```
Sub AddHyperlink() 
 Selection.TextRange.Hyperlinks.Add Text:=Selection.TextRange, _ 
 Address:="http://www.tailspintoys.com/" 
End Sub
```

Использование **гиперссылок** (индекс), где индекс — номер индекса, для возврата объекта **гиперссылки** в публикации, диапазон или выбора. В этом примере отображает адрес гиперссылки, если указанный выделенный фрагмент содержит гиперссылки.




```
Sub DisplayHyperlinkAddress() 
 With Selection.TextRange.Hyperlinks 
 If .Count > 0 Then _ 
 MsgBox .Item(1).Address 
 End With 
End Sub
```

Свойство **[Count](http://msdn.microsoft.com/library/36747f3e-b365-11ca-9cbe-f6148f7da235%28Office.15%29.aspx)** для данного семейства сайтов возвращает число гиперссылки в указанные форму или только выделенного фрагмента.


## <a name="methods"></a>Методы



|**Name**|
|:-----|
|[Добавление](http://msdn.microsoft.com/library/f5a8cc01-a571-623d-bfab-fe48e43a21b1%28Office.15%29.aspx)|

## <a name="properties"></a>Properties



|**Name**|
|:-----|
|[Приложения](http://msdn.microsoft.com/library/c025e261-dc0e-9445-2c89-c9e79db6b3cd%28Office.15%29.aspx)|
|[Count](http://msdn.microsoft.com/library/36747f3e-b365-11ca-9cbe-f6148f7da235%28Office.15%29.aspx)|
|[Элемент](http://msdn.microsoft.com/library/8d288fc6-9ded-5732-b972-6fa366ef31c3%28Office.15%29.aspx)|
|[Родительский раздел](http://msdn.microsoft.com/library/e3b25f19-6322-172a-3620-c3e728074655%28Office.15%29.aspx)|

