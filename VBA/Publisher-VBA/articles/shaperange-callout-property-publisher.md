---
title: "Свойство ShapeRange.Callout (издатель)"
keywords: vbapb10.chm2293811
f1_keywords: vbapb10.chm2293811
ms.prod: publisher
api_name: Publisher.ShapeRange.Callout
ms.assetid: 25b9b444-6cbf-085a-df7f-8899e8e55057
ms.date: 06/08/2017
ms.openlocfilehash: 297b13511e711a6535a3044f7f33afc177e0be3b
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="shaperangecallout-property-publisher"></a>Свойство ShapeRange.Callout (издатель)

Возвращает объект **[CalloutFormat](calloutformat-object-publisher.md)** , представляющий форматирование выноски строки.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Выноски**

 переменная _expression_A, представляющий объект **ShapeRange** .


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


