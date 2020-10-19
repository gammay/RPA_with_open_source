*** Settings ***
Library    ExcelUtil
Library    FinUtil
Library    DateTime

*** Variables ***
${file}  D:\\holdings.xlsx

*** Tasks ***
Networth Calculator
    @{holdings}=  Get Data  ${file}
    FOR  ${holding}  IN  @{holdings}
        ${row}  Set Variable  ${holding[0]}
        ${company}  Set Variable  ${holding[1]}
        ${code}  Set Variable  ${holding[2]}
        ${num}  Set Variable  ${holding[3]}
        ${price}=  Get Quote  ${code}
        ${dt}=  Get Time  NOW
        ${dt_rowcol}=  Set Variable  D${row}
        Set Data  ${file}  ${dt_rowcol}  ${dt}
        ${price_rowcol}=  Set Variable  E${row}
        Set Data  ${file}  ${price_rowcol}  ${price}
    END
