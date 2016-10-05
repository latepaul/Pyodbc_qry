#!/bin/sh
if [ -z "$II_SYSTEM" ]
then
    echo "II_SYSTEM needs to be set!"
    exit 1
fi

export ODBCSYSINI=$II_SYSTEM/ingres/files
export ODBCINI=$II_SYSTEM/ingres/files/odbc.ini
export II_ODBC_WCHAR_SIZE=2

python pyodbc_qry.py
