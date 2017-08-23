---
title: "Свойство CalloutFormat.Border (издатель)"
keywords: vbapb10.chm2490628
f1_keywords: vbapb10.chm2490628
ms.prod: publisher
api_name: Publisher.CalloutFormat.Border
ms.assetid: 64a72ec7-4cc8-f0c7-9858-45e97bac0411
ms.date: 06/08/2017
ms.openlocfilehash: 6e37a6456836853f6b4910ebdd35e483814748e7
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="calloutformatborder-property-publisher"></a>Свойство CalloutFormat.Border (издатель)

Возвращает или задает константой **MsoTriState**, указывающее, является ли текст в указанном выноски заключается в границы. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Границы**

 переменная _expression_A, представляет собой объект- **CalloutFormat** .


### <a name="return-value"></a>Возвращаемое значение

MsoTriState


## <a name="remarks"></a>Заметки

Значение свойства **границы** может иметь одно из ** [MsoTriState](http://msdn.microsoft.com/library/2036cfc9-be7d-e05c-bec7-af05e3c3c515%28Office.15%29.aspx)** объявленные константы в библиотеке типов, Microsoft Office.


## <a name="example"></a>Пример

В этом примере добавляется овала active публикации и выноски, указывающий на овал. Текст выноски будут иметь границы, но не вертикальная черта, отделяющий текст из строки выноски.


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
 .Accent = msoFalse 
 .Border = msoTrue 
 End With 
 End With 
End With 

```


