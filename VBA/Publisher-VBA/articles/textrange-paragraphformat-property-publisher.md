---
title: "Свойство TextRange.ParagraphFormat (издатель)"
keywords: vbapb10.chm5308439
f1_keywords: vbapb10.chm5308439
ms.prod: publisher
api_name: Publisher.TextRange.ParagraphFormat
ms.assetid: 475da411-9292-a12d-addd-1bbe822ec09e
ms.date: 06/08/2017
ms.openlocfilehash: 57308aa2f14ec7b26506e7f700e97e1479c15e95
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="textrangeparagraphformat-property-publisher"></a>Свойство TextRange.ParagraphFormat (издатель)

Возвращает объект **[ParagraphFormat](paragraphformat-object-publisher.md)** , представляющий форматирование абзаца для указанного текста диапазон или стиля текста.


## <a name="syntax"></a>Синтаксис

 _выражение_. **ParagraphFormat**

 переменная _expression_A, представляющий объект **TextRange** .


## <a name="example"></a>Пример

Следующий пример удаляет все табуляции из текста в первую фигуру на странице один из активных публикации.


```vb
Dim pfTemp As ParagraphFormat 
 
Set pfTemp = ActiveDocument.Pages(1).Shapes(1) _ 
 .TextFrame.TextRange.ParagraphFormat 
 
pfTemp.Tabs.ClearAll
```


