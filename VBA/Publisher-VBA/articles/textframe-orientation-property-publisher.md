---
title: "Свойство TextFrame.Orientation (издатель)"
keywords: vbapb10.chm3866659
f1_keywords: vbapb10.chm3866659
ms.prod: publisher
api_name: Publisher.TextFrame.Orientation
ms.assetid: f510e624-6322-4054-5e7f-8688c5ea817a
ms.date: 06/08/2017
ms.openlocfilehash: 86712e54f57b2bbda7bd48632ecf920e09994bd5
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="textframeorientation-property-publisher"></a>Свойство TextFrame.Orientation (издатель)

Возвращает или задает значение константы **PbTextOrientation**, представляющий потоки текст в текстовом поле. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Ориентация**

 переменная _expression_A, представляет собой объект- **TextFrame** .


### <a name="return-value"></a>Возвращаемое значение

PbTextOrientation


## <a name="remarks"></a>Заметки

Значение свойства **ориентации** может иметь одно из **[PbTextOrientation](pbtextorientation-enumeration-publisher.md)** константы в библиотеке типов, Microsoft Publisher.


## <a name="example"></a>Пример

В этом примере задается Ориентация текста в поле указанный текст для вертикальной так, чтобы текст перетекал сверху вниз. Предполагается, что имеется по крайней мере один фигуры на странице один активный публикации.






```vb
Sub SetVerticalTextBox() 
 ActiveDocument.Pages(1).Shapes(1).TextFrame _ 
 .Orientation = pbTextOrientationVerticalEastAsia 
End Sub
```


