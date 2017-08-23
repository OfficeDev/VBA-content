---
title: "Свойство Font.ExpandUsingKashida (издатель)"
keywords: vbapb10.chm5374004
f1_keywords: vbapb10.chm5374004
ms.prod: publisher
api_name: Publisher.Font.ExpandUsingKashida
ms.assetid: ecf3a170-5f07-379e-ff56-504beb770308
ms.date: 06/08/2017
ms.openlocfilehash: c8ad0ee66f85381f13c84dac67b41d537aa32258
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="fontexpandusingkashida-property-publisher"></a>Свойство Font.ExpandUsingKashida (издатель)

Возвращает или задает константой **MsoTriState** , указывающее, следует ли применять правил кашиды при применении отслеживания арабский текст. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **ExpandUsingKashida**

 переменная _expression_A, представляющий объект **шрифта** .


### <a name="return-value"></a>Возвращаемое значение

MsoTriState


## <a name="remarks"></a>Заметки

Значение свойства **ExpandUsingKashida** может иметь одно из **MsoTriState** константы объявляются в библиотеке типов Microsoft Office и показаны в следующей таблице.



|**Константы**|**Описание**|
|:-----|:-----|
| **msoFalse**| Microsoft Publisher не применяется правил кашиды при применении отслеживания арабский текст.|
| **msoTriStateMixed**|Возвращает значение, указывающее, сочетание **msoTrue** и **msoFalse** для диапазона указанный текст.|
| **msoTriStateToggle**|Задайте значение, могут переключаться между **msoTrue** и **msoFalse**.|
| **msoTrue**| Publisher применения правил кашиды при применении отслеживания арабский текст.|

## <a name="example"></a>Пример

В следующем примере задается Publisher для применения правил кашиды при применении отслеживания арабского текста для всех диапазонов текст на странице один из активных публикации.


```vb
Dim shpLoop As Shape 
 
For Each shpLoop In ActiveDocument.Pages(1).Shapes 
 If shpLoop.HasTextFrame Then 
 shpLoop.TextFrame.TextRange _ 
 .Font.ExpandUsingKashida = msoTrue 
 End If 
Next shpLoop
```


