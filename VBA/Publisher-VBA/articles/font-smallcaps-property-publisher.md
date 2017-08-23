---
title: "Свойство Font.SmallCaps (издатель)"
keywords: vbapb10.chm5373971
f1_keywords: vbapb10.chm5373971
ms.prod: publisher
api_name: Publisher.Font.SmallCaps
ms.assetid: ab50b850-f371-7d8e-0c19-00ad68e700f0
ms.date: 06/08/2017
ms.openlocfilehash: a81e63916ef42ac7b4c5002a0f83779d944658b6
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="fontsmallcaps-property-publisher"></a>Свойство Font.SmallCaps (издатель)

Возвращает или задает константой **MsoTriState** , указывающее, указанный текст в формате малые прописные буквы. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **SmallCaps**

 переменная _expression_A, представляющий объект **Font** .


### <a name="return-value"></a>Возвращаемое значение

MsoTriState


## <a name="remarks"></a>Заметки

Значение свойства **SmallCaps** может иметь одно из **MsoTriState** константы объявляются в библиотеке типов Microsoft Office и показаны в следующей таблице.



|**Константы**|**Описание**|
|:-----|:-----|
| **msoFalse**|Ни один из символов в диапазоне форматируются как малые прописные буквы.|
| **msoTriStateMixed**|Возвращает значение, указывающее, сочетание **msoTrue** и **msoFalse** для диапазона указанной фигуры.|
| **msoTriStateToggle**|Задайте значение, могут переключаться между **msoTrue** и **msoFalse**.|
| **msoTrue**| Все символы в диапазоне форматируются как малые прописные буквы.|
Установка для свойства **SmallCaps** **msoTrue** удаляет все прописные диапазон текста.


## <a name="example"></a>Пример

В этом примере проверяется текст во второй материал и, если он имеет смешанный форматирование, формат весь текст на малые прописные буквы.


```vb
Sub SmallCap() 
 
 Dim fntSC As Font 
 
 Set fntSC = Application.ActiveDocument.Stories(2).TextRange.Font 
 With fntSC 
 If .SmallCaps = msoTriStateMixed Then 
 .SmallCaps = msoTrue 
 Else 
 MsgBox "Mixed small caps are not in this story." 
 End If 
 End With 
 
End Sub
```


