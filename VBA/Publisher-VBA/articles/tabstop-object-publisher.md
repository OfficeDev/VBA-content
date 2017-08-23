---
title: "Объект TabStop (издатель)"
keywords: vbapb10.chm5701631
f1_keywords: vbapb10.chm5701631
ms.prod: publisher
api_name: Publisher.TabStop
ms.assetid: 74e71d75-503f-ef57-ddeb-24a788402df2
ms.date: 06/08/2017
ms.openlocfilehash: e05ca924ce5c63324612ee5e3bf7e188f6aa2013
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="tabstop-object-publisher"></a>Объект TabStop (издатель)

Представляет один позиции табуляции. Объект **TabStop** является элементом коллекции **[TabStops](tabstops-object-publisher.md)** . Коллекция **TabStops** представляет все пользовательские и табуляции по умолчанию в абзаца или группы абзацев.
 


## <a name="remarks"></a>Заметки

Присвойте свойству **[DefaultTabStop](document-defaulttabstop-property-publisher.md)** интервал табуляции по умолчанию.
 

 

## <a name="example"></a>Пример

Используйте **[вкладки](tabstops-add-method-publisher.md)** (индекс), где индекс — это расположение табуляции (в пунктах) или номер индекса, чтобы возвратить объект **TabStop** . Табуляции, индексируются числовым слева направо на линейке. Следующий пример удаляет первый настраиваемых табуляции из выделенных абзацев.
 

 

```
Sub ClearTabStop() 
 Selection.TextRange.ParagraphFormat.Tabs(1).Clear 
End Sub
```

В следующем примере добавляется по правому краю позиции табуляции, размещенный в 2 дюйма для выделенных абзацев.
 

 



```
Sub ChangeTabStop() 
 Selection.TextRange.ParagraphFormat.Tabs(2) _ 
 .Alignment = pbTabAlignmentTrailing 
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


## <a name="methods"></a>Методы



|**Name**|
|:-----|
|[Очистить](tabstop-clear-method-publisher.md)|

## <a name="properties"></a>Properties



|**Name**|
|:-----|
|[Выравнивание](tabstop-alignment-property-publisher.md)|
|[Приложения](tabstop-application-property-publisher.md)|
|[Ведущий сотрудник](tabstop-leader-property-publisher.md)|
|[Родительский раздел](tabstop-parent-property-publisher.md)|
|[Position](tabstop-position-property-publisher.md)|

