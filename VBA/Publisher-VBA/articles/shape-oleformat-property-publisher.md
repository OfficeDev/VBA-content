---
title: "Свойство Shape.OLEFormat (издатель)"
keywords: vbapb10.chm2228327
f1_keywords: vbapb10.chm2228327
ms.prod: publisher
api_name: Publisher.Shape.OLEFormat
ms.assetid: 36bffb6b-4c7b-85f9-87b3-d7d7c1aed134
ms.date: 06/08/2017
ms.openlocfilehash: afc409e1b9a676293d9bd52de886babca454933e
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="shapeoleformat-property-publisher"></a>Свойство Shape.OLEFormat (издатель)

Возвращает объект **[OLEFormat](oleformat-object-publisher.md)** , который содержит параметры форматирования для указанного фигуры OLE. Применяется к **фигуры** или **ShapeRange** объектов, представляющих объекты OLE.


## <a name="syntax"></a>Синтаксис

 _выражение_. **OLEFormat**

 переменная _expression_A, представляющий объект **фигуры** .


## <a name="example"></a>Пример

В этом примере циклически просматривает все фигуры на первой странице активного документа и автоматически обновляет все связанные листы Excel.


```vb
Sub UpdateLinkedExcelSpreadsheets() 
 Dim shp As Shape 
 For Each shp In ActiveDocument.Pages(1).Shapes 
 If shp.Type = msoLinkedOLEObject Then 
 If shp.OLEFormat.ProgId = "Excel.Sheet" Then 
 shp.LinkFormat.Update 
 End If 
 End If 
 Next shp 
End Sub
```


