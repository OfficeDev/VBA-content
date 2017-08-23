---
title: "Свойство Font.ItalicBi (издатель)"
keywords: vbapb10.chm5373969
f1_keywords: vbapb10.chm5373969
ms.prod: publisher
api_name: Publisher.Font.ItalicBi
ms.assetid: 604e776c-92b0-6e5b-2599-ab879c61a78a
ms.date: 06/08/2017
ms.openlocfilehash: afc1f35e4ab4ce8199391721e9698cbb384d7f47
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="fontitalicbi-property-publisher"></a>Свойство Font.ItalicBi (издатель)

Возвращает или задает константой **MsoTriState** , указывающее, указанный текст в формате italic; применяется для текста справа налево языке. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **ItalicBi**

 переменная _expression_A, представляющий объект **шрифта** .


### <a name="return-value"></a>Возвращаемое значение

MsoTriState


## <a name="remarks"></a>Заметки

Значение свойства **ItalicBi** может иметь одно из **MsoTriState** константы объявляются в библиотеке типов Microsoft Office и показаны в следующей таблице.



|**Константы**|**Описание**|
|:-----|:-----|
| **msoFalse**| Ни один из символов в диапазоне форматируются как курсив.|
| **msoTriStateMixed**|Возвращает значение, указывающее, сочетание **msoTrue** и **msoFalse** для диапазона указанной фигуры.|
| **msoTriStateToggle**|Задайте значение, могут переключаться между **msoTrue** и **msoFalse**.|
| **msoTrue**|Все символы в диапазоне форматируются как курсив.|

## <a name="example"></a>Пример

В этом примере проверяется текст в первый сценариев и отображается одно из двух возможных текстовых полей, в зависимости от того, является ли текста справа налево отформатированные и ли шрифт представлен в формате курсив.


```vb
Sub ItalicRtoL() 
 
 Dim stf As Font 
 
 Set stf = Application.ActiveDocument.Stories(1).TextRange.Font 
 With stf 
 If .ItalicBi = msoTrue Then 
 MsgBox "This story is right-to-left and is formatted as italic." 
 Else 
 MsgBox "This story is either not right-to-left" &; _ 
 " or it is not formatted as italic" 
 End If 
 End With 
 
End Sub
```


