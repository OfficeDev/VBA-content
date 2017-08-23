---
title: "Свойство Shape.WebTextBox (издатель)"
keywords: vbapb10.chm2228342
f1_keywords: vbapb10.chm2228342
ms.prod: publisher
api_name: Publisher.Shape.WebTextBox
ms.assetid: 8a3f8389-728f-b8ae-3c89-dc8d03a3818e
ms.date: 06/08/2017
ms.openlocfilehash: ae697e84d71e730fab6bb1c648444b57117608f3
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="shapewebtextbox-property-publisher"></a>Свойство Shape.WebTextBox (издатель)

Возвращает объект **[WebTextBox](webtextbox-object-publisher.md)** , связанный с указанным фигуры.


## <a name="syntax"></a>Синтаксис

 _выражение_. **WebTextBox**

 переменная _expression_A, представляющий объект **фигуры** .


### <a name="return-value"></a>Возвращаемое значение

WebTextBox


## <a name="example"></a>Пример

В этом примере создается новое текстовое поле Web, задает текст по умолчанию, указывает, что запись является обязательным и ограничения для записи до 50 символов.


```vb
Dim shpNew As Shape 
Dim wtbTemp As WebTextBox 
 
Set shpNew = ActiveDocument.Pages(1).Shapes _ 
 .AddWebControl(Type:=pbWebControlSingleLineTextBox, _ 
 Left:=100, Top:=100, Width:=150, Height:=15) 
 
Set wtbTemp = shpNew.WebTextBox 
 
With wtbTemp 
.DefaultText = "Please Enter Your Full Name" 
 .RequiredControl = msoTrue 
 .Limit = 50 
End With
```


