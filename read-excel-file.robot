*** Settings ***
Resource        ./resource.robot

*** Test Cases ***

01-Read-and-compare-xlsx
    Compare Excel
    ...    C:\\Desktop\\Excel\\file-old.xlsx
    ...    C:\\Desktop\\Excel\\file-new.xlsx