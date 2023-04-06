#!/bin/bash

file=`ls -t ~/Загрузки/RU_INTERNAL_RECOVERY*.xlsx | head -1`
xlsx2csv -s5 $file > /tmp/recovery_pdl_zero.csv
xlsx2csv -s8 $file > /tmp/recovery_pdl_repeated.csv
xlsx2csv -s3 $file > /tmp/recovery_mini.csv
xlsx2csv -s4 $file > /tmp/recovery_mini56.csv
./recovery_checker.py > /tmp/recovery.html
firefox /tmp/recovery.html
