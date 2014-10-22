//
//  ExcelDates.cpp
//  QC-FXOptions
//
//  Created by Alvaro Diaz on 15-11-12.
//  Copyright (c) 2012 Alvaro Diaz. All rights reserved.
//

#include "ExcelDates.h"

void ExcelSerialDateToDMY(int nSerialDate, int &nDay, int &nMonth, int &nYear)
{
    // Excel/Lotus 123 have a bug with 29-02-1900. 1900 is not a
    // leap year, but Excel/Lotus 123 think it is...
    if (nSerialDate == 60)
    {
        nDay    = 29;
        nMonth    = 2;
        nYear    = 1900;
        
        return;
    }
    else if (nSerialDate < 60)
    {
        // Because of the 29-02-1900 bug, any serial date
        // under 60 is one off... Compensate.
        nSerialDate++;
    }
    
    // Modified Julian to DMY calculation with an addition of 2415019
    int l = nSerialDate + 68569 + 2415019;
    int n = int(( 4 * l ) / 146097);
    l = l - int(( 146097 * n + 3 ) / 4);
    int i = int(( 4000 * ( l + 1 ) ) / 1461001);
    l = l - int(( 1461 * i ) / 4) + 31;
    int j = int(( 80 * l ) / 2447);
    nDay = l - int(( 2447 * j ) / 80);
    l = int(j / 11);
    nMonth = j + 2 - ( 12 * l );
    nYear = 100 * ( n - 49 ) + i + l;
}

int DMYToExcelSerialDate(int nDay, int nMonth, int nYear)
{
    // Excel/Lotus 123 have a bug with 29-02-1900. 1900 is not a
    // leap year, but Excel/Lotus 123 think it is...
    if (nDay == 29 && nMonth == 02 && nYear==1900)
        return 60;
    
    // DMY to Modified Julian calculate with an extra substraction of 2415019.
    long nSerialDate =
    int(( 1461 * ( nYear + 4800 + int(( nMonth - 14 ) / 12) ) ) / 4) +
    int(( 367 * ( nMonth - 2 - 12 * ( ( nMonth - 14 ) / 12 ) ) ) / 12) -
    int(( 3 * ( int(( nYear + 4900 + int(( nMonth - 14 ) / 12) ) / 100) ) ) / 4) +
    nDay - 2415019 - 32075;
    
    if (nSerialDate <= 60)
    {
        //Aqui hay que poner <=60
        // Because of the 29-02-1900 bug, any serial date
        // under 60 is one off... Compensate.
        nSerialDate--;
    }
    
    return (int)nSerialDate;
}

int dayOfWeekFromDate(int y, int m, int d)
{
    static int t[] = {0, 3, 2, 5, 0, 3, 5, 1, 4, 6, 2, 4};
    y -= m < 3;
    return (y + y/4 - y/100 + y/400 + t[m-1] + d) % 7;
}

void businessDay(int &nDay, int &nMonth, int &nYear)
{
    int day = dayOfWeekFromDate(nYear, nMonth, nDay);
    int serial = DMYToExcelSerialDate(nDay, nMonth, nYear);
    
    if (day == 0)
    {
        serial = serial + 1;
        ExcelSerialDateToDMY(serial, nDay, nMonth, nYear);
        return;
    }
    if (day == 6)
    {
        serial = serial + 2;
        ExcelSerialDateToDMY(serial, nDay, nMonth, nYear);
        return;
    }
    
    return;
}

