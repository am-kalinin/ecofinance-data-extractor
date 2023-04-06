#!/bin/bash

file=`ls -t ../Загрузки/csv*.csv | head -1`
./simple_mess.py $file > messengers.html
firefox messengers.html
