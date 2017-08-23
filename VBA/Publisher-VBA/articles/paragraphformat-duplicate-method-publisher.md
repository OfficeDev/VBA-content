---
title: "Метод ParagraphFormat.Duplicate (издатель)"
keywords: vbapb10.chm5439510
f1_keywords: vbapb10.chm5439510
ms.prod: publisher
api_name: Publisher.ParagraphFormat.Duplicate
ms.assetid: 83156999-7867-05c2-9e85-4cc0f580ac6e
ms.date: 06/08/2017
ms.openlocfilehash: 12048fac340edba4d1f21e77723093cd74846393
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="paragraphformatduplicate-method-publisher"></a>Метод ParagraphFormat.Duplicate (издатель)

Создает копию на указанный объект **[ParagraphFormat](paragraphformat-object-publisher.md)** и возвращает новый объект **ParagraphFormat** .


## <a name="syntax"></a>Синтаксис

 _выражение_. **Дублирующиеся**

 переменная _expression_A, представляет собой объект- **ParagraphFormat** .


### <a name="return-value"></a>Возвращаемое значение

ParagraphFormat


## <a name="example"></a>Пример

В следующем примере дублирует форматирование информации из диапазона текст в фигуре одно на странице абзаца, один из активных публикации и применяется к диапазон текста в форме двух.


```vb
Dim pfTemp As ParagraphFormat 
 
With ActiveDocument.Pages(1) 
 Set pfTemp = .Shapes(1).TextFrame _ 
 .TextRange.ParagraphFormat.Duplicate 
 .Shapes(2).TextFrame _ 
 .TextRange.ParagraphFormat = pfTemp 
End With
```


