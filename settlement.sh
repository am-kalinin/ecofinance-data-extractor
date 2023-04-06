#!/bin/bash

file=`ls -t ~/Загрузки/Prolongation_term*.xlsx | head -1`
xlsx2csv $file > /tmp/settlement.csv
./valid_settlement_loans.py > /tmp/settlement.html
firefox /tmp/settlement.html
