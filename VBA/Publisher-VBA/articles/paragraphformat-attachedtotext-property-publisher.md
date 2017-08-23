---
title: "Свойство ParagraphFormat.AttachedToText (издатель)"
keywords: vbapb10.chm5439512
f1_keywords: vbapb10.chm5439512
ms.prod: publisher
api_name: Publisher.ParagraphFormat.AttachedToText
ms.assetid: 1bfb902c-d728-1f97-513c-dcee54ce57a8
ms.date: 06/08/2017
ms.openlocfilehash: 6d23054d82bc9eaa48d85cab25fbbf4f1ff8303d
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="paragraphformatattachedtotext-property-publisher"></a>Свойство ParagraphFormat.AttachedToText (издатель)

 **Значение true,** Если объект **шрифта** или **ParagraphFormat** присоединен к объекту **TextRange** . Если объект связан объект **TextRange** , документ будет обновляться при изменении свойства объекта. Если объект не подключена, в документе будет изменяться, пока объект применяется объект **TextRange** или **стиля** . Только для чтения **типа Boolean**.


## <a name="syntax"></a>Синтаксис

 _выражение_. **AttachedToText**

 переменная _expression_A, представляет собой объект- **ParagraphFormat** .


## <a name="example"></a>Пример

В этом примере дублирует форматирование; шрифта Затем проверяется ли дублируемые форматирования подключенный к диапазон текста, и если он не установлен, подключает форматирования для второй фигуры.


```vb
Sub DuplicateText() 
 Dim fntTemp As Font 
 With ActiveDocument.Pages(1) 
 Set fntTemp = .Shapes(1).TextFrame.TextRange.Font.Duplicate 
 If fntTemp.AttachedToText <> True Then _ 
 ActiveDocument.Pages(1).Shapes(2) _ 
 .TextFrame.TextRange.Font = fntTemp 
 End With 
End Sub
```


