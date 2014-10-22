//
//  ExcelDates.h
//  QC-FXOptions
//
//  Created by Alvaro Diaz on 15-11-12.
//  Copyright (c) 2012 Alvaro Diaz. All rights reserved.
//

#ifndef __QC_FXOptions__ExcelDates__
#define __QC_FXOptions__ExcelDates__

#include <iostream>

#endif /* defined(__QC_FXOptions__ExcelDates__) */

void ExcelSerialDateToDMY(int nSerialDate, int &nDay,
                          int &nMonth, int &nYear);

int DMYToExcelSerialDate(int nDay, int nMonth, int nYear);

int dayOfWeekFromDate(int y, int m, int d);

// void businessDay(int &nDay, int &nMonth, int &nYear);

