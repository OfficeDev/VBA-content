---
title: "Свойство OLEFormat.ObjectVerbs (издатель)"
keywords: vbapb10.chm4456453
f1_keywords: vbapb10.chm4456453
ms.prod: publisher
api_name: Publisher.OLEFormat.ObjectVerbs
ms.assetid: 887070e6-7f7d-4f65-290e-3d46bfd91d34
ms.date: 06/08/2017
ms.openlocfilehash: 368a1832013d3d2acea0abb681a71f1ddee04e4c
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="oleformatobjectverbs-property-publisher"></a>Свойство OLEFormat.ObjectVerbs (издатель)

Возвращает коллекцию **[ObjectVerbs](objectverbs-object-publisher.md)** , содержащий все команды OLE для указанного объекта OLE. Только для чтения.


## <a name="syntax"></a>Синтаксис

 _выражение_. **ObjectVerbs**

 переменная _expression_A, представляющий объект **OLEFormat** .


### <a name="return-value"></a>Возвращаемое значение

ObjectVerbs


## <a name="example"></a>Пример

В этом примере отображаются все доступные команды для объекта OLE, содержащихся в форму одно на вторую страницу в активной публикации. В данном примере для работы фигуры один должен быть фигуры, представляющий объект OLE.


```vb
Dim v As String 
 
With ActiveDocument.Pages(2).Shapes(1).OLEFormat 
 For Each v In .ObjectVerbs 
 MsgBox v 
 Next 
End With
```


