---
title: "Метод Shapes.AddCatalogMergeFieldToCanvas (издатель)"
keywords: vbapb10.chm2162760
f1_keywords: vbapb10.chm2162760
ms.prod: publisher
api_name: Publisher.Shapes.AddCatalogMergeFieldToCanvas
ms.assetid: 30cd45d0-97f0-ab01-31c2-8d819b435b1b
ms.date: 06/08/2017
ms.openlocfilehash: 85613151a60cdae10beb13fe9367f67aa5a7fe23
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# Метод Shapes.AddCatalogMergeFieldToCanvas (издатель)

Добавляет поле слияния каталога указанного типа на основе. Возвращает значение nothing.


## Синтаксис

 _выражение_. **AddCatalogMergeFieldToCanvas** ( **_CanvasId_**, **_CatalogMergeFieldType_**, **_DbCol_**)

 переменная _expression_A, представляет собой объект- **фигур** .


### Параметры



|**Имя**|**Обязательный или необязательный**|**Тип данных**|**Описание**|
|:-----|:-----|:-----|:-----|
|CanvasId|Обязательное свойство.| **[INT]**|Идентификатор полотно, в которую нужно добавить поле слияния каталога.|
|CatalogMergeFieldType|Обязательное свойство.| **pbCatalogMergeFieldType**|Тип (рисунок или текст) поле слияния каталога для добавления.|
|DbCol|Обязательное свойство.| **[INT]**|Количество столбцов в источнике данных, который содержит сведения объединения каталога.|

