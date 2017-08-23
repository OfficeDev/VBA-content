---
title: "Свойство CalloutFormat.Accent (издатель)"
keywords: vbapb10.chm2490624
f1_keywords: vbapb10.chm2490624
ms.prod: publisher
api_name: Publisher.CalloutFormat.Accent
ms.assetid: 8e31544c-79ed-3882-98d1-42fc88f58115
ms.date: 06/08/2017
ms.openlocfilehash: d310973a1bbf767ec864335482e255602d131ab8
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="calloutformataccent-property-publisher"></a>Свойство CalloutFormat.Accent (издатель)

Возвращает или задает константой **MsoTriState**, указывающее, является ли вертикальная черта разделяет выноски текста из строки выноски. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Диакритических знаков**

 переменная _expression_A, представляет собой объект- **CalloutFormat** .


### <a name="return-value"></a>Возвращаемое значение

MsoTriState


## <a name="remarks"></a>Заметки

Значение свойства **Акцент** может иметь одно из следующих констант **MsoTriState** .



|**Константы**|**Описание**|
|:-----|:-----|
| **msoCTrue**|Не используется с этим свойством.|
| **msoFalse**|Вертикальная черта отделяйте выноски текст из строки выноски.|
| **msoTriStateMixed**|Возвращаемое значение. Указывает сочетание **msoTrue** и **msoFalse** в диапазоне указанные форму.|
| **msoTriStateToggle**|Заданное значение. Переключение между **msoTrue** и **msoFalse**.|
| **msoTrue**|Вертикальная черта разделяет текст выноски линии выноски.|

## <a name="example"></a>Пример

В этом примере добавляется овала active публикации и выноски, указывающий на овал. Текст выноски не будут иметь границы, но он будет иметь вертикальная черта, отделяющий текст из строки выноски.


```vb
With ActiveDocument.Pages(1).Shapes 
 ' Add an oval. 
 .AddShape Type:=msoShapeOval, _ 
 Left:=180, Top:=200, Width:=280, Height:=130 
 
 ' Add a callout. 
 With .AddCallout(Type:=msoCalloutTwo, _ 
 Left:=420, Top:=170, Width:=170, Height:=40) 
 
 ' Add text to the callout. 
 .TextFrame.TextRange.Text = "This is an oval" 
 
 ' Add an accent bar to the callout. 
 With .Callout 
 .Accent = msoTrue 
 .Border = msoFalse 
 End With 
 End With 
End With 

```


