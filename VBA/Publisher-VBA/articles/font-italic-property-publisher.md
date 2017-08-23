---
title: "Свойство Font.Italic (издатель)"
keywords: vbapb10.chm5373968
f1_keywords: vbapb10.chm5373968
ms.prod: publisher
api_name: Publisher.Font.Italic
ms.assetid: c55c0bfa-a365-86ac-4cfb-f6911dadd0af
ms.date: 06/08/2017
ms.openlocfilehash: 48f34ec79a7fef25c2019a7b2774b628b40c70b3
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="fontitalic-property-publisher"></a>Свойство Font.Italic (издатель)

Возвращает или задает константой **MsoTriState** , указывающее, указанный текст в формате курсив. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Курсив**

 переменная _expression_A, представляющий объект **шрифта** .


### <a name="return-value"></a>Возвращаемое значение

MsoTriState


## <a name="remarks"></a>Заметки

Значение свойства **Italic** может иметь одно из **MsoTriState** константы объявляются в библиотеке типов Microsoft Office и показаны в следующей таблице.



|**Константы**|**Описание**|
|:-----|:-----|
| **msoFalse**|Ни один из символов в диапазоне форматируются как курсив.|
| **msoTriStateMixed**|Возвращает значение, указывающее, сочетание **msoTrue** и **msoFalse** для диапазона указанной фигуры.|
| **msoTriStateToggle**|Задайте значение, могут переключаться между **msoTrue** и **msoFalse**.|
| **msoTrue**|Все символы в диапазоне форматируются как курсив.|

## <a name="example"></a>Пример

В этом примере проверяется весь текст во второй материал active публикации и, если он имеет текст в формате курсив, задает весь текст на курсив. Если текст не все курсивом или все курсивом, отображается сообщение о том, что не не смешанных курсивное форматирование.


```vb
Sub ItalicStory() 
 
 Dim stf As Font 
 
 Set stf = Application.ActiveDocument.Stories(2).TextRange.Font 
 With stf 
 If .Italic = msoTriStateMixed Then 
 .Italic = msoTrue 
 Else 
 MsgBox "There is no mixed italic formatting in this story." 
 End If 
 End With 
 
End Sub
```


