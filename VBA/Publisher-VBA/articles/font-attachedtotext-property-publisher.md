---
title: "Свойство Font.AttachedToText (издатель)"
keywords: vbapb10.chm5373989
f1_keywords: vbapb10.chm5373989
ms.prod: publisher
api_name: Publisher.Font.AttachedToText
ms.assetid: 23b0519a-9f35-fa25-752a-4942e8161edd
ms.date: 06/08/2017
ms.openlocfilehash: d547d99e5adaff66fba8ef69d1fc0960899cfc2f
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="fontattachedtotext-property-publisher"></a>Свойство Font.AttachedToText (издатель)

 **Значение true,** Если объект **шрифта** или **ParagraphFormat** присоединен к объекту **TextRange** . Если объект связан объект **TextRange** , документ будет обновляться при изменении свойства объекта. Если объект не подключена, в документе будет изменяться, пока объект применяется объект **TextRange** или **стиля** . Только для чтения **типа Boolean**.


## <a name="syntax"></a>Синтаксис

 _выражение_. **AttachedToText**

 переменная _expression_A, представляющий объект **Font** .


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


