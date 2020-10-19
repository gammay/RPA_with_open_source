#!/bin/sh

# set -x

file="holdings.csv"

networth=0

while IFS= read -r line
do
    company=`echo "$line" | cut -d',' -f1`
    if [ "$company" = "COMPANY" ]
    then
        continue
    fi

    if [ "$company" = "" ]
    then
        break
    fi

    code=`echo $line | cut -d',' -f2`

    num=`echo $line | cut -d',' -f3`

    price=`curl https://query1.finance.yahoo.com/v7/finance/options/$code 2>/dev/null | grep -o '"regularMarketPrice":[0-9\.]*' | cut -d':' -f2`
    value=`echo ${price}*${num} | bc`
    echo "$company ($code): $value ($price * $num)"

    networth=`echo ${networth}+${value} | bc`

done < $file

echo "Networth on `date` is $networth"

echo "\"`date`\",$networth" >> $file

