---
title: "Свойство Font.Bold (издатель)"
keywords: vbapb10.chm5373955
f1_keywords: vbapb10.chm5373955
ms.prod: publisher
api_name: Publisher.Font.Bold
ms.assetid: 3b9ba2b0-c319-9d08-9a36-5b292046962e
ms.date: 06/08/2017
ms.openlocfilehash: cef00dc0647605ae053348888102ab92f7b03a23
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="fontbold-property-publisher"></a>Свойство Font.Bold (издатель)

Возвращает или задает константой **MsoTriState**, представляющее состояние свойства **Bold** символов в диапазон текста. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Полужирным шрифтом**

 переменная _expression_A, представляющий объект **Font** .


### <a name="return-value"></a>Возвращаемое значение

MsoTriState


## <a name="remarks"></a>Заметки

Значение свойства **Bold** может иметь одно из следующих **MsoTriState** константы, описанные в библиотеке типов, Microsoft Office.



|**Константы**|**Описание**|
|:-----|:-----|
| **msoFalse**|Ни один из символов в диапазоне форматируются полужирным шрифтом.|
| **msoTriStateMixed**|Возвращает значение, указывающее, что диапазон содержит текст полужирным и не форматированный текст полужирным шрифтом.|
| **msoTriStateToggle**|Задайте значение, могут переключаться между **msoTrue** и **msoFalse**.|
| **msoTrue**|Все символы в диапазоне форматируются полужирным шрифтом.|

## <a name="example"></a>Пример

В этом примере проверяется весь текст во второй материал active публикации и, если она содержит текст полужирным шрифтом и не полужирным шрифтом, задает весь текст выделяется полужирным шрифтом. Если текст все полужирным или не все полужирным шрифтом, отображается сообщение, предупреждающее, что нет не смешанных полужирное начертание. Для этого кода для выполнения должным образом должен быть более двух функциональности с текстом в активной публикации.


```vb
Sub BoldStory() 
 
 Dim stf As Publisher.Font 
 
 Set stf = Application.ActiveDocument.Stories(2).TextRange.Font 
 With stf 
 If .Bold = msoTriStateMixed Then 
 .Bold = msoTrue 
 Else 
 MsgBox "Mixed bolding is not in this story." 
 End If 
 End With 
 
End Sub
```


