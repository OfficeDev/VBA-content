---
title: "Свойство WebTextBox.RequiredControl (издатель)"
keywords: vbapb10.chm4194310
f1_keywords: vbapb10.chm4194310
ms.prod: publisher
api_name: Publisher.WebTextBox.RequiredControl
ms.assetid: 32e18d4b-7af0-b079-4baf-9acc07c3c37d
ms.date: 06/08/2017
ms.openlocfilehash: c8eadf4d573765f657a4747050fd4628ec98af86
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="webtextboxrequiredcontrol-property-publisher"></a>Свойство WebTextBox.RequiredControl (издатель)

Указывает, необходима ли запись в элемент управления текстового поля Web. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **RequiredControl**

 переменная _expression_A, представляет собой объект- **WebTextBox** .


### <a name="return-value"></a>Возвращаемое значение

MsoTriState


## <a name="remarks"></a>Заметки

Значение свойства **RequiredControl** может иметь одно из **MsoTriState** константы объявляются в библиотеке типов Microsoft Office и показаны в следующей таблице.



|**Константы**|**Описание**|
|:-----|:-----|
| **msoFalse**|Указывает, что запись в указанном Web текстового поля не является обязательным.|
| **msoTrue**| Указывает, что запись в указанном Web текстового поля является обязательным.|

## <a name="example"></a>Пример

В этом примере создается новый элемент управления полем Web текст в активной публикации, задает текст по умолчанию и ограничение количества знаков для текстового поля и указывает, что запись является обязательным.


```vb
Sub AddWebTextBoxControl() 
 With ActiveDocument.Pages(1).Shapes.AddWebControl _ 
 (Type:=pbWebControlMultiLineTextBox, Left:=72, _ 
 Top:=72, Width:=300, Height:=100).WebTextBox 
 .DefaultText = "Please enter text here." 
 .Limit = 200 
 .RequiredControl = msoTrue 
 End With 
End Sub
```


