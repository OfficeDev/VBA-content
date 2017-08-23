---
title: "Объект TabStops (издатель)"
keywords: vbapb10.chm5636095
f1_keywords: vbapb10.chm5636095
ms.prod: publisher
api_name: Publisher.TabStops
ms.assetid: fbaa194c-754a-3437-c3d5-fa70c951ca4f
ms.date: 06/08/2017
ms.openlocfilehash: 01f0e8a70e8ff28e6b9bef381c378fd8cf4c4887
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="tabstops-object-publisher"></a>Объект TabStops (издатель)

Коллекция объектов **[TabStop](tabstop-object-publisher.md)** , представляющие пользовательские и по умолчанию вкладки для абзаца или группы абзацев.
 


## <a name="example"></a>Пример

Используйте свойство **[вкладки](paragraphformat-tabs-property-publisher.md)** для возврата коллекции **TabStops** . В следующем примере удаляются все, что настраиваемые табуляции из первого абзаца в активной публикации.
 

 

```
Sub ClearAllTabStops() 
 ActiveDocument.Pages(1).Shapes(1).TextFrame.TextRange _ 
 .ParagraphFormat.Tabs.ClearAll 
End Sub
```

В следующем примере добавляет позиции табуляции, размещенный в 2,5 дюйма для выделенных абзацев и затем отображает положение каждого элемента в коллекции **TabStops** .
 

 



```
Sub Tabs() 
 Dim intTab As Integer 
 Selection.TextRange.ParagraphFormat.Tabs _ 
 .Add Position:=InchesToPoints(2.5), _ 
 Alignment:=pbTabAlignmentLeading, Leader:=pbTabLeaderNone 
 With Selection.TextRange.ParagraphFormat 
 For intTab = 1 To .Tabs.Count 
 MsgBox "Position = " &amp; PointsToInches _ 
 (.Tabs(intTab).Position) &amp; " inches" 
 intTab = intTab + 1 
 Next intTab 
 End With 
End Sub
```

Используйте метод **[Add](tabstops-add-method-publisher.md)** для добавления позиции табуляции. В следующем примере добавляется два табуляции для выделенных абзацев. Первый табуляции — это вкладка по левому краю с точками заполнитель, размещенный в 1 дюйм (72 точки). Второй позиции табуляции выравнивается по центру и размещенный в 2 дюйма.
 

 



```
Sub AddNewTabs() 
 With Selection.TextRange.ParagraphFormat.Tabs 
 .Add Position:=InchesToPoints(1), _ 
 Leader:=pbTabLeaderDot, Alignment:=pbTabAlignmentLeading 
 .Add Position:=InchesToPoints(2), _ 
 Leader:=pbTabLeaderNone, Alignment:=pbTabAlignmentCenter 
 End With 
End Sub
```

Используйте **[вкладки](paragraphformat-tabs-property-publisher.md)** (индекс), где индекс — это расположение табуляции (в пунктах) или номер индекса, чтобы возвратить объект **TabStop** . Табуляции, индексируются числовым слева направо на линейке. Следующий пример удаляет первый stop пользовательской вкладки в первый абзац в активной публикации.
 

 



```
Sub ClearTabStop() 
 ActiveDocument.Pages(1).Shapes(1).TextFrame.TextRange _ 
 .ParagraphFormat.Tabs(1).Clear 
End Sub
```

В следующем примере изменяется второй вкладки в выделении до табуляции по правому краю.
 

 



```
Sub ChangeTabStop() 
 Selection.TextRange.ParagraphFormat.Tabs(2) _ 
 .Alignment = pbTabAlignmentTrailing 
End Sub
```


## <a name="methods"></a>Методы



|**Name**|
|:-----|
|[Добавление](tabstops-add-method-publisher.md)|
|[ClearAll](tabstops-clearall-method-publisher.md)|
|[Элемент](tabstops-item-method-publisher.md)|

## <a name="properties"></a>Properties



|**Name**|
|:-----|
|[Приложения](tabstops-application-property-publisher.md)|
|[Count](tabstops-count-property-publisher.md)|
|[Родительский раздел](tabstops-parent-property-publisher.md)|

