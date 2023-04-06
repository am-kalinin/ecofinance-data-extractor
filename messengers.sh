#!/bin/bash

file=`ls -t ~/Загрузки/csv*.* | head -1`
./mess_maker.py $file > /tmp/messengers.html
firefox /tmp/messengers.html
