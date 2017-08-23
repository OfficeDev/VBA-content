---
title: "Свойство TextEffectFormat.KernedPairs (издатель)"
keywords: vbapb10.chm3735813
f1_keywords: vbapb10.chm3735813
ms.prod: publisher
api_name: Publisher.TextEffectFormat.KernedPairs
ms.assetid: 1382ae7a-250f-ca08-a57f-f7132078e3f2
ms.date: 06/08/2017
ms.openlocfilehash: bc32131017e9cdec1918088c8863ee794bd81b40
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="texteffectformatkernedpairs-property-publisher"></a>Свойство TextEffectFormat.KernedPairs (издатель)

Задает или возвращает константу **MsoTriState** , которое указывает, ли кернинг пар знаков в объекте WordArt. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **KernedPairs**

 переменная _expression_A, представляет собой объект- **TextEffectFormat** .


### <a name="return-value"></a>Возвращаемое значение

MsoTriState


## <a name="remarks"></a>Заметки

Значение свойства **KernedPairs** может иметь одно из **MsoTriState** константы объявляются в библиотеке типов Microsoft Office и показаны в следующей таблице.



|**Константы**|**Описание**|
|:-----|:-----|
| **msoFalse**| Пары символов в указанном объекте WordArt не имеют были кернинг.|
| **msoTriStateToggle**|Переключение между **msoTrue** и **msoFalse**.|
| **msoTrue**|Ли кернинг пар знаков в указанном объекте WordArt.|

## <a name="example"></a>Пример

В этом примере включается Кернинг знаков для всех объектов WordArt в активной публикации.


```vb
Sub KernedWordArt() 
 Dim pagPage As Page 
 Dim shpShape As Shape 
 For Each pagPage In ActiveDocument.Pages 
 For Each shpShape In pagPage.Shapes 
 If shpShape.Type = msoTextEffect Then 
 shpShape.TextEffect.KernedPairs = msoTrue 
 End If 
 Next 
 Next 
End Sub
```


