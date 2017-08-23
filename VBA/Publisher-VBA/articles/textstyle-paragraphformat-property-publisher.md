---
title: "Свойство TextStyle.ParagraphFormat (издатель)"
keywords: vbapb10.chm5963781
f1_keywords: vbapb10.chm5963781
ms.prod: publisher
api_name: Publisher.TextStyle.ParagraphFormat
ms.assetid: 5ab0a2ec-d7a9-f3af-29e7-5421427ee783
ms.date: 06/08/2017
ms.openlocfilehash: 76cba87ca96ae0c59d9126577b3c772be642e90d
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="textstyleparagraphformat-property-publisher"></a>Свойство TextStyle.ParagraphFormat (издатель)

Возвращает объект **[ParagraphFormat](paragraphformat-object-publisher.md)** , представляющий форматирование абзаца для указанного текста диапазон или стиля текста.


## <a name="syntax"></a>Синтаксис

 _выражение_. **ParagraphFormat**

 переменная _expression_A, представляющий объект **стиля текста** .


## <a name="example"></a>Пример

Следующий пример удаляет все табуляции из текста в первую фигуру на странице один из активных публикации.


```vb
Dim pfTemp As ParagraphFormat 
 
Set pfTemp = ActiveDocument.Pages(1).Shapes(1) _ 
 .TextFrame.TextRange.ParagraphFormat 
 
pfTemp.Tabs.ClearAll
```


