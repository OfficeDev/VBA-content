---
title: "Свойство Shape.Callout (издатель)"
keywords: vbapb10.chm2228275
f1_keywords: vbapb10.chm2228275
ms.prod: publisher
api_name: Publisher.Shape.Callout
ms.assetid: e0682bb4-1129-fa58-b28c-46d7ce2fad0c
ms.date: 06/08/2017
ms.openlocfilehash: 4ca97c958ac94fb522f7be9c697deec32c1650ca
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="shapecallout-property-publisher"></a>Свойство Shape.Callout (издатель)

Возвращает объект **[CalloutFormat](calloutformat-object-publisher.md)** , представляющий форматирование выноски строки.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Выноски**

 переменная _expression_A, представляющий объект **фигуры** .


## <a name="example"></a>Пример

В этом примере добавляется овала active публикации и выноски, указывающий на овал. Текст выноски не будут иметь границу, но он будет иметь вертикальная черта, отделяющий текст из строки выноски.


```vb
Sub NewShapeItem() 
 
 Dim shpNew As Shapes 
 
 Set shpNew = Application.ActiveDocument.MasterPages(1).Shapes 
 With shpNew 
 .AddShape Type:=msoShapeOval, Left:=180, _ 
 Top:=200, Width:=280, Height:=130 
 With .AddCallout(Type:=msoCalloutTwo, Left:=420, _ 
 Top:=170, Width:=170, Height:=40) 
 .TextFrame.TextRange = "Big Oval" 
 With .Callout 
 .Accent = msoTrue 
 .Border = msoFalse 
 End With 
 End With 
 End With 
 
End Sub
```


