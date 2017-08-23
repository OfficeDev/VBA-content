---
title: "Свойство WebTextBox.DefaultText (издатель)"
keywords: vbapb10.chm4194307
f1_keywords: vbapb10.chm4194307
ms.prod: publisher
api_name: Publisher.WebTextBox.DefaultText
ms.assetid: 348c1bc2-61c9-f89f-5e7a-b73ddaa3d216
ms.date: 06/08/2017
ms.openlocfilehash: d0f82bc05557caf4cb3ee05660525b2decf24e15
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="webtextboxdefaulttext-property-publisher"></a>Свойство WebTextBox.DefaultText (издатель)

Возвращает или задает **строку** , представляющую текст по умолчанию в элемент управления текстового поля Web. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **DefaultText**

 переменная _expression_A, представляет собой объект- **WebTextBox** .


### <a name="return-value"></a>Возвращаемое значение

String


## <a name="example"></a>Пример

В этом примере создается новый элемент управления полем Web текст в активной публикации, задает текст по умолчанию и ограничение количества знаков для текстового поля и указывает, что необходимый элемент управления.


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


