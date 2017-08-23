---
title: "Свойство ShapeRange.OLEFormat (издатель)"
keywords: vbapb10.chm2293863
f1_keywords: vbapb10.chm2293863
ms.prod: publisher
api_name: Publisher.ShapeRange.OLEFormat
ms.assetid: 237b51e8-dced-3e21-d257-410121107a63
ms.date: 06/08/2017
ms.openlocfilehash: 54d15ca195118b28446df0f6a68d1fccc5823802
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="shaperangeoleformat-property-publisher"></a>Свойство ShapeRange.OLEFormat (издатель)

Возвращает объект **[OLEFormat](oleformat-object-publisher.md)** , который содержит параметры форматирования для указанного фигуры OLE. Применяется к **фигуры** или **ShapeRange** объектов, представляющих объекты OLE.


## <a name="syntax"></a>Синтаксис

 _выражение_. **OLEFormat**

 переменная _expression_A, представляющий объект **ShapeRange** .


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


