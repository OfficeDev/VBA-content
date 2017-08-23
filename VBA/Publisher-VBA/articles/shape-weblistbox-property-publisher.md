---
title: "Свойство Shape.WebListBox (издатель)"
keywords: vbapb10.chm2228341
f1_keywords: vbapb10.chm2228341
ms.prod: publisher
api_name: Publisher.Shape.WebListBox
ms.assetid: c100dfc7-6fbd-db48-4de9-4a9a49739a8f
ms.date: 06/08/2017
ms.openlocfilehash: 149ced2b126146d9dacbcf0d34e82bad4b270e41
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="shapeweblistbox-property-publisher"></a>Свойство Shape.WebListBox (издатель)

Возвращает объект **[WebListBox](weblistbox-object-publisher.md)** , связанный с указанным фигуры.


## <a name="syntax"></a>Синтаксис

 _выражение_. **WebListBox**

 переменная _expression_A, представляющий объект **фигуры** .


### <a name="return-value"></a>Возвращаемое значение

WebListBox


## <a name="example"></a>Пример

В этом примере создается новый список Web и добавляет несколько элементов. Обратите внимание, что при создании, элемент управления списка Web содержит три элемента по умолчанию. В этом примере включает в себя цикл, который удаляет поле элементов списка по умолчанию, прежде чем добавлять новые элементы.


```vb
Dim shpNew As Shape 
Dim wlbTemp As WebListBox 
Dim intCount As Integer 
 
Set shpNew = ActiveDocument.Pages(1).Shapes _ 
 .AddWebControl(Type:=pbWebControlListBox, Left:=100, _ 
 Top:=150, Width:=300, Height:=72) 
 
Set wlbTemp = shpNew.Web ListBox 
 
With wlbTemp 
 .MultiSelect = msoFalse 
 
 With .ListBoxItems 
 For intCount = 1 To .Count 
 .Delete (1) 
 Next intCount 
 
 .AddItem Item:="Green" 
 .AddItem Item:="Purple" 
 .AddItem Item:="Red" 
 .AddItem Item:="Black" 
 End With 
End With
```


