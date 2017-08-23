---
title: "Свойство Font.SuperScript (издатель)"
keywords: vbapb10.chm5373972
f1_keywords: vbapb10.chm5373972
ms.prod: publisher
api_name: Publisher.Font.SuperScript
ms.assetid: 582c02c9-4dcb-f826-8ec0-e9e10702f717
ms.date: 06/08/2017
ms.openlocfilehash: 2361adfc905ff6604f709a943fc36815caedbde9
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="fontsuperscript-property-publisher"></a>Свойство Font.SuperScript (издатель)

Возвращает или задает константой **MsoTriState** , указывающее, форматируются ли символы как надстрочный знак в диапазоне указанный текст. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Надстрочный знак**

 переменная _expression_A, представляющий объект **Font** .


### <a name="return-value"></a>Возвращаемое значение

MsoTriState


## <a name="remarks"></a>Заметки

Значение свойства **надстрочный знак** может быть одной из констант **MsoTriState** объявлена в библиотеке типов, Microsoft Office и показаны в следующей таблице.



|**Константы**|**Описание**|
|:-----|:-----|
| **msoFalse**| Нет символов в диапазоне форматируются как надстрочный знак.|
| **msoTriStateMixed**|Возвращает значение, указывающее, сочетание **msoTrue** и **msoFalse** для диапазона указанной фигуры.|
| **msoTriStateToggle**|Задайте значение, могут переключаться между **msoTrue** и **msoFalse**.|
| **msoTrue**|Все символы в диапазоне форматируются как надстрочный знак.|
Установка для свойства **надстрочный знак** **msoTrue** удаляет нижнего индекса форматирование из диапазона текста.


## <a name="example"></a>Пример

В этом примере проверяется текст во второй материал и, если он имеет смешанный надстрочное начертание, он форматирует весь текст как надстрочный знак.


```vb
Sub SuperScript() 
 
 Dim fntSuper As Font 
 
 Set fntSuper = Application.ActiveDocument.Stories(2).TextRange.Font 
 With fntSuper 
 If .SuperScript = msoTriStateMixed Then 
 .SuperScript = msoTrue 
 Else 
 MsgBox "Mixed superscript not in this story." 
 End If 
 End With 
 
End Sub
```


