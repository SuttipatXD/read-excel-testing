*** Settings ***
Library     SeleniumLibrary
Library     String
Library     ExcelUtil
Library     ExcelLibrary


*** Keywords ***
Compare Excel
    [Arguments]    ${excel-file-old}    ${excel-file-new}
    Open Excel    ${excel-file-old}
    Log    [1]PATH FILE: ${excel-file-old}
    ${column-count-old}=    Get Column Count    Sheet1
    Log    [1]TOTAL COLUMN: ${column-count-old}
    FOR    ${index-column}    IN RANGE    1    ${column-count-old}
        ${data}=    Read Cell Data By Coordinates    Sheet1    1    ${index-column}
        Log    Cell at Row 1, Column ${index-column} and data: ${data}
    END
    Close All Excel Documents

    Open Excel    ${excel-file-new}
    Log    [2]PATH FILE: ${excel-file-new}
    ${column-count-new}=    Get Column Count    Sheet1
    Log    [2]TOTAL COLUMN: ${column-count-new}
    Close All Excel Documents
    FOR    ${index-column}    IN RANGE    1    ${column-count-new}
        ${data}=    Read Cell Data By Coordinates    Sheet1    1    ${index-column}
        Log    Cell at Row 1, Column ${index-column} and data: ${data}
    END

    IF    '${column-count-old}' == '${column-count-new}'
        Log    Excel files have the SAME number of columns
        # เปรียบเทียบข้อมูลในคอลัมน์
        FOR    ${index-column}    IN RANGE    1    ${column-count-old}
            ${data-old}=    Read Cell Data By Coordinates    Sheet1    1    ${index-column}
            ${data-new}=    Read Cell Data By Coordinates    Sheet1    1    ${index-column}
            IF    '${data-old}' == '${data-new}'
                Log    Data in old file and new file are the same in column: ${index-column}
            ELSE
                Log    Data in old file and new file are different in column: ${index-column}
            END
        END
    ELSE
        Log    Excel files have DIFFERENT number of columns
    END
