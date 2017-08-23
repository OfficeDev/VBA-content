---
title: "Свойство ParagraphFormat.Tabs (издатель)"
keywords: vbapb10.chm5439506
f1_keywords: vbapb10.chm5439506
ms.prod: publisher
api_name: Publisher.ParagraphFormat.Tabs
ms.assetid: c42ba898-b84f-7215-129d-8134670f75ac
ms.date: 06/08/2017
ms.openlocfilehash: ddc2249ee8e02a4859950b05f2ad119310840882
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="paragraphformattabs-property-publisher"></a>Свойство ParagraphFormat.Tabs (издатель)

Возвращает **[TabStops](tabstops-object-publisher.md)** объект, представляющий пользовательские и по умолчанию вкладки для абзаца или группы абзацев.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Вкладки**

 переменная _expression_A, представляет собой объект- **ParagraphFormat** .


### <a name="return-value"></a>Возвращаемое значение

TabStops


## <a name="example"></a>Пример

В следующем примере добавляется два табуляции для выделенных абзацев. Первый табуляции — это вкладка по левому краю с точками заполнитель, размещенный в 1 дюйм (72 точки). Второй позиции табуляции выравнивается по центру и размещенный в 2 дюйма.


```vb
Dim tabsAll As TabStops 
 
Set tabsAll = Selection.TextRange.ParagraphFormat.Tabs 
 
With tabsAll 
 .Add Position:=InchesToPoints(1), _ 
 Leader:=pbTabLeaderDot, Alignment:=pbTabAlignmentLeading 
 .Add Position:=InchesToPoints(2), _ 
 Leader:=pbTabLeaderNone, Alignment:=pbTabAlignmentCenter 
End With
```


