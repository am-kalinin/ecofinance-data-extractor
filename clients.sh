#!/bin/bash

file=`ls -t ~/Загрузки/RU_1-120DPD*.* | head -1`
#echo $file
xlsx2csv $file > /tmp/clients.csv
file=`ls -t ~/Загрузки/csv*.csv | head -1`
./late_loans_check.py $file > /tmp/clients.html
firefox /tmp/clients.html
firefox /tmp/mini56.html
