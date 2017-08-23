---
title: "Свойство Font.SubScript (издатель)"
keywords: vbapb10.chm5373973
f1_keywords: vbapb10.chm5373973
ms.prod: publisher
api_name: Publisher.Font.SubScript
ms.assetid: 9992fdcc-dd60-b2f7-307b-99b10dc7debb
ms.date: 06/08/2017
ms.openlocfilehash: 23500ec814c9075702f55691aa7d0eeee070b4f5
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="fontsubscript-property-publisher"></a>Свойство Font.SubScript (издатель)

Возвращает или задает константой **MsoTriState** , указывающее, форматируются ли символы как подстрочный знак в диапазоне указанный текст. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Подстрочный знак**

 переменная _expression_A, представляющий объект **Font** .


### <a name="return-value"></a>Возвращаемое значение

MsoTriState


## <a name="remarks"></a>Заметки

Значение свойства **подстрочный знак** может быть одной из констант **MsoTriState** объявлена в библиотеке типов, Microsoft Office и показаны в следующей таблице.



|**Константы**|**Описание**|
|:-----|:-----|
| **msoFalse**|Нет символов в диапазоне форматируются как индекс.|
| **msoTriStateMixed**|Возвращает значение, указывающее, сочетание **msoTrue** и **msoFalse** для диапазона указанной фигуры.|
| **msoTriStateToggle**|Задайте значение, могут переключаться между **msoTrue** и **msoFalse**.|
| **msoTrue**| Все символы в диапазоне форматируются как индекс.|
Установка для свойства **подстрочный знак** **msoTrue** удаляет верхним форматирования из диапазона текста.


## <a name="example"></a>Пример

В этом примере проверяется текст во второй материал и, если он имеет смешанный нижней индексации, он форматирует весь текст как индекс.


```vb
Sub SubScript() 
 
 Dim fntSS As Font 
 
 Set fntSS = Application.ActiveDocument.Stories(2).TextRange.Font 
 With fntSS 
 If .SubScript = msoTriStateMixed Then 
 .SubScript = msoTrue 
 Else 
 MsgBox "Mixed subscript not in this story." 
 End If 
 End With 
 
End Sub
```


