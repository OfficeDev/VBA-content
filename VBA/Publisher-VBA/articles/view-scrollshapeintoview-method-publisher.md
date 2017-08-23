---
title: "Метод View.ScrollShapeIntoView (издатель)"
keywords: vbapb10.chm327685
f1_keywords: vbapb10.chm327685
ms.prod: publisher
api_name: Publisher.View.ScrollShapeIntoView
ms.assetid: 1d654fd4-d3b8-49e4-731d-fed27e6e0d8d
ms.date: 06/08/2017
ms.openlocfilehash: 8670a38b8fea306a31aebb990d973f8e9885a2ee
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="viewscrollshapeintoview-method-publisher"></a>Метод View.ScrollShapeIntoView (издатель)

Прокрутка окна публикации для отображения указанного фигуры в окне Публикация или области.


## <a name="syntax"></a>Синтаксис

 _выражение_. **ScrollShapeIntoView** ( **_Фигуры_**)

 переменная _expression_A, представляющий объект **View** .


### <a name="parameters"></a>Параметры



|**Имя**|**Обязательный или необязательный**|**Тип данных**|**Описание**|
|:-----|:-----|:-----|:-----|
|Shape|Обязательное свойство.| **Фигура**|Фигура появится в представлении.|

## <a name="example"></a>Пример

В этом примере добавляется фигура на новую страницу и прокрутка текущего представления на новую фигуру.


```vb
Sub ScrollIntoView() 
 Dim shpStar As Shape 
 Dim intWidth As Integer 
 Dim intHeight As Integer 
 
 With ActiveDocument 
 intWidth = .PageSetup.PageWidth 
 intWidth = (intWidth / 2) - 75 
 intHeight = .PageSetup.PageHeight 
 intHeight = (intHeight / 2) - 75 
 
 With .Pages.Add(Count:=1, After:=ActiveDocument.Pages.Count) 
 Set shpStar = .Shapes.AddShape(Type:=msoShape5pointStar, _ 
 Left:=intWidth, Top:=intHeight, Width:=150, Height:=150) 
 shpStar.TextFrame.TextRange.Text = "New Star Shape" 
 End With 
 End With 
 
 ActiveView.ScrollShapeIntoView Shape:=shpStar 
 
End Sub
```


