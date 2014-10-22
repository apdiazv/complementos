#include "QCDate.h"
#include "ExcelDates.h"

#include <iostream>
#include <cassert>
#include <vector>
#include <cmath>

////////////////////////////////////////////////////////////////
// Constructors
////////////////////////////////////////////////////////////////

// Overridden default constructor
QCDate::QCDate()
{
	_serialDate = 0;
	_day = 0;
	_month = 1;
	_year = 1900;
	_dayOfWeek = 6;

}

// Overridden copy constructor
// Copies entries of other Date into it
QCDate::QCDate(const QCDate& otherDate)
{
	_serialDate = otherDate._serialDate;
	_day = otherDate._day;
	_month = otherDate._month;
	_year = otherDate._year;
	_dayOfWeek = otherDate._dayOfWeek;
}

// Constructor for date of a given excel serial number
QCDate::QCDate(int excelSerialDate)
{
	assert(excelSerialDate > 0);
	_serialDate = excelSerialDate;
	ExcelSerialDateToDMY(_serialDate, _day, _month, _year);
	_dayOfWeek = dayOfWeekFromDate(_year, _month, _day);
}

//	Constructor for date of a given a day, month and year
QCDate::QCDate(int day, int month, int year)
{
	_serialDate = DMYToExcelSerialDate(day, month, year);
	ExcelSerialDateToDMY(_serialDate, _day, _month, _year);
	_dayOfWeek = dayOfWeekFromDate(_year, _month, _day);
}

////////////////////////////////////////////////////////////////
// Destructor
////////////////////////////////////////////////////////////////
QCDate::~QCDate(void)
{
}


////////////////////////////////////////////////////////////////
// Methos
////////////////////////////////////////////////////////////////

// Method to get the serial number
int QCDate::excelSerialDate() const
{
	return _serialDate;
}

// Method to get the day of a date
int QCDate::day() const
{
	return _day;
}

// Method to get the month of a date
int QCDate::month() const
{
	return _month;
}

// Method to get the year of a date
int QCDate::year() const
{
	return _year;
}

// Method to get the day of week of a date
int QCDate::dayOfWeek()
{
	return _dayOfWeek;
}

// Method to add or subtract days to a date to calculate a past or future date
QCDate QCDate::addDays(int days)
{
	int newSerialDate = _serialDate + days;
	assert(newSerialDate > 0);
	QCDate date(newSerialDate);
	return date;
}

// Method to add or subtract month from a given date
QCDate QCDate::addMonths(int months)
{
	std::vector<int> daysMonths(12);
	std::vector<int>::iterator it;
	int newDay = _day;
	int newMonth = _month;
	int newYear = _year;
	bool eom = false;

	int aux[] = {31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31};
	daysMonths.assign(aux, aux + 12);
	it = daysMonths.begin()+ newMonth - 1;

	// detect if the date correspond to the end of a month
	if (newDay == *it &&  _month != 2)
	{
		eom = true;
	}

	months = months + newMonth;
	newMonth = months - (int)( 12 * ( months / 12));
	if (newMonth == 0)
	{
		newMonth = 12;
	}

	newYear = newYear + (int)(( months - 1) / 12 );
    if (newYear < 1900)
	{
		newYear = 1900;
	}

	it = daysMonths.begin()+ newMonth - 1;
	if (newDay > *it)
	{
		eom = true;
	}

    // if the date correspond to the end of a month
	if (eom)
	{
		newDay = *it; // the resulting date is forced to be the end of a month
		if(newMonth == 2 && (newYear % 4 == 0) && (newYear % 100 != 0 )||(newYear % 400 == 0 ))
		{
			newDay = 29;
		}
	}

	QCDate date(newDay, newMonth, newYear);

	return date;
}


////////////////////////////////////////////////////////////////
// Operators
////////////////////////////////////////////////////////////////

// Overloading the assignment operator
QCDate& QCDate::operator =(const QCDate& otherDate)
{
	assert(otherDate._serialDate > 0);
	_serialDate = otherDate._serialDate;
	_day = otherDate._day;
	_month = otherDate._month;
	_year = otherDate._year;
	_dayOfWeek = otherDate._dayOfWeek;

	return *this;
}
// Overloading the post-increment operator ++
QCDate QCDate::operator++(int)
{
	int newSerialDate = _serialDate +1;
	QCDate date(newSerialDate);

	return date;
}

// Overloading the binary + operator
QCDate QCDate::operator +(const QCDate& date1) const
{
	assert(date1._serialDate > 0);
	int newSerialDate = _serialDate + date1._serialDate;
	QCDate date(newSerialDate);

	return date;
}

// Overloading the binary - operator
int QCDate::operator -(const QCDate& date1) const
{
	return _serialDate - date1._serialDate;
}

// Overloading the binary < operator
bool QCDate::operator <(const QCDate& date1) const
{
	if(_serialDate < date1._serialDate)
	{
		return true;
	}
	return false;
}

// Overloading the binary > operator
bool QCDate::operator >(const QCDate& date1) const
{
	if(_serialDate > date1._serialDate)
	{
		return true;
	}

	return false;
}

////////////////////////////////////////////////////////////////
// Friends functions
////////////////////////////////////////////////////////////////

// Friend style function to get the business day (only exclude Weekends)
QCDate getBussDay(const QCDate& date1)
{
    int day = date1._dayOfWeek;
    int serial = date1._serialDate;
    
    if (day == 0)
    {
        serial = serial + 1;
        // ExcelSerialDateToDMY(serial, _day, _month, _year);
		QCDate date(serial);

        return date;
    }
    if (day == 6)
    {
        serial = serial + 2;
        // ExcelSerialDateToDMY(serial, _day, _month, _year);
		QCDate date(serial);

        return date;
    }

    QCDate date(serial);

    return date;
}

// Friend style function to get the previous business day
QCDate getPrevDay(const QCDate& date1)
{
    int day = date1._dayOfWeek;
    int serial = date1._serialDate;
    
    if (day == 0)
    {
        serial = serial - 2;
        // ExcelSerialDateToDMY(serial, _day, _month, _year);
		QCDate date(serial);
        return date;
    }
    if (day == 6)
    {
        serial = serial - 1;
        // ExcelSerialDateToDMY(serial, _day, _month, _year);
		QCDate date(serial);
        return date;
    }

    QCDate date(serial);

    return date;
}