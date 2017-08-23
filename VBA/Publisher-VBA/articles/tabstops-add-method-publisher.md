---
title: "Метод TabStops.Add (издатель)"
keywords: vbapb10.chm5570565
f1_keywords: vbapb10.chm5570565
ms.prod: publisher
api_name: Publisher.TabStops.Add
ms.assetid: 23536810-e851-c0ac-22e2-fab41582d612
ms.date: 06/08/2017
ms.openlocfilehash: 061585c5545dabb2008dbdba419200d9a86f4f6d
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="tabstopsadd-method-publisher"></a>Метод TabStops.Add (издатель)

Добавление новой позиции табуляции определенной коллекции **TabStops** .


## <a name="syntax"></a>Синтаксис

 _выражение_. **Добавление** ( **_Положение_**, **_Выравнивание_**, **_ведущий_**)

 переменная _expression_A, представляет собой объект- **TabStops** .


### <a name="parameters"></a>Параметры



|**Имя**|**Обязательный или необязательный**|**Тип данных**|**Описание**|
|:-----|:-----|:-----|:-----|
|Position|Обязательное свойство.| **Variant**|Горизонтальную позицию новой позиции табуляции относительно левого края рамки. Числовые значения вычисляются в точках; строки вычисляются в единицах, заданных и может быть в любой единицы измерения, поддерживаются в Microsoft Publisher (например, «2,5 дюйма»).|
|Выравнивание|Обязательное свойство.| **PbTabAlignmentType**|Настройка выравнивания для позиции табуляции.|
|Ведущий сотрудник|Обязательное свойство.| **PbTabLeaderType**|Тип ведущий табуляции.|

## <a name="remarks"></a>Заметки

Выравнивание может иметь одно из следующих констант PbTabAlignmentType.



| **pbTabAlignmentCenter**|| **pbTabAlignmentDecimal**|| **pbTabAlignmentLeading**|| **pbTabAlignmentTrailing**| Ведущий может иметь одно из следующих констант **PbTabLeaderType** .



| **pbTabLeaderBullet**|| **pbTabLeaderDashes**|| **pbTabLeaderDot**|| **pbTabLeaderLine**|| **pbTabLeaderNone**|

## <a name="example"></a>Пример

В следующем примере добавляется новый по левому краю табуляции 0,5 дюйма от левого края кадра указанный текст.


```vb
ActiveDocument.Pages(1).Shapes(1).TextFrame _ 
 .TextRange.ParagraphFormat.Tabs _ 
 .Add Position:="0.5 in", _ 
 Alignment:=pbTabAlignmentLeading, _ 
 Leader:=pbTabLeaderNone
```


