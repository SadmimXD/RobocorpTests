*** Comments ***
    Open Workbook    C:\\Users\\Sadmim Hossain\\Downloads\\orders.csv
    ${robot_Orders}=    Read Worksheet As Table    header=True
    Close Workbook
    FOR    ${order_Details}    IN    @{robot_Orders}
    Send Order in website    ${order_Details}
    END

Send Order in website
    [Arguments]    ${order_Details}
    Select From List By Value    head    ${order_Details}[Head]

IF    ${ErrorPop}
    Log    Error Found
    Log    ${ErrorPop}
    ${ErrorMessage}=    Get Text    css:#root > div > div.container > div > div.col-sm-7 > div
    Log    Error Pop Up Found ${ErrorMessage}
    Click Button    order
    ELSE
    Log    ${ErrorPop}
    ${ErrorMessage}=    Get Text    id:receipt
    Log    Error Pop Up Found ${ErrorMessage}
    END


*** Settings ***
Documentation       Orders robots from RobotSpareBin Industries Inc.
...                 Saves the order HTML receipt as a PDF file.
...                 Saves the screenshot of the ordered robot.
...                 Embeds the screenshot of the robot to the PDF receipt.
...                 Creates ZIP archive of the receipts and the images.

Library             RPA.Browser.Selenium    auto_close=${FALSE}
Library             RPA.HTTP
Library             RPA.Excel.Files
Library             RPA.Tables
Library             Screenshot
Library             RPA.Excel.Application
Resource            ../Course1/tasks.robot
Library             RPA.PDF
Library             RPA.Desktop
Library             RPA.FileSystem
Library             RPA.Archive


*** Variables ***
${csvFileName}=     robotorder.csv
${download_Dir}=    SecondCourse/output/Robots
${csvURL}=          https://robotsparebinindustries.com/orders.csv


*** Tasks ***
Order robots from RobotSpearBin Industries Inc
    Open the robot order website

Download csv
    Download csv file and read it
    Zip all the files
    [Teardown]    Close the Browser


*** Keywords ***
Open the robot order website
    Open Available Browser    https://robotsparebinindustries.com/#/robot-order
    Maximize Browser Window

Click Okay pop Up
    Click Button    css:#root > div > div.modal > div > div > div > div > div > button.btn.btn-dark

Download csv file and read it
    Download    ${csvURL}    overwrite=True
    Read table from CSV    orders.csv
    ${table}=    Read table from CSV    orders.csv
    Close Workbook
    FOR    ${order_Details}    IN    @{table}
        Send Order in website    ${order_Details}
    END

Send Order in website
    [Arguments]    ${order_Details}
    Click Okay pop Up
    Select From List By Value    head    ${order_Details}[Head]
    Click Element
    ...    css:#root > div > div.container > div > div.col-sm-7 > form > div:nth-child(2) > div > div:nth-child(${order_Details}[Body]) > label
    Input Text    Xpath:/html/body/div/div/div[1]/div/div[1]/form/div[3]/input    ${order_Details}[Legs]
    Input Text    address    ${order_Details}[Address]
    Click Button    preview
    Sleep    1s
    Wait Until Element Is Visible    robot-preview-image
    Click Button    order
    Error PopUp
    Embading Screenshot to PDF    ${order_Details}[Order number]
    Click Button    order-another

Take Screenshot
    [Arguments]    ${filename}
    Screenshot    id:robot-preview-image    ${OUTPUT_DIR}${/}${download_Dir}${/}robot${filename}.png

Save pdf and submit
    [Arguments]    ${filename}
    Wait Until Element Is Visible    robot-preview-image
    ${receipt}=    Get Element Attribute    id:receipt    outerHTML
    Html To Pdf    ${receipt}    ${download_Dir}${/}robot${filename}.pdf

    # ${screenshot}=    Take Screenshot    ${row}[Order number]

Error PopUp
    ${ErrorPop}=    Is Element Visible    id:receipt
    IF    ${ErrorPop}
        Log    Receipt found
    ELSE
        Click Button    order
        ${ErrorPop}=    Is Element Visible    order
        WHILE    ${ErrorPop}
            Sleep    1s
            Click Button    order
            ${ErrorPop}=    Is Element Visible    order
        END
    END

Close the Browser
    Close Browser

Embading Screenshot to PDF
    [Arguments]    ${filename}
    Take Screenshot    ${filename}
    Save pdf and submit    ${filename}
    ${pdf}=    Open Pdf    ${download_Dir}${/}robot${filename}.pdf
    ${files}=    Create List    ${download_Dir}${/}robot${filename}.png    ${download_Dir}${/}robot${filename}.pdf
    #${ss}=    ${download_Dir}${/}robot${filename}.png
    Add Files To Pdf    ${files}    ${download_Dir}${/}robot${filename}.pdf
    Close Pdf    ${pdf}
    Remove File    ${download_Dir}${/}robot${filename}.png

Zip all the files
    ${allfiles}=    List Files In Directory    ${OUTPUT_DIR}${/}${download_Dir}
    FOR    ${file}    IN    @{allfiles}
        Log    ${file}
    END
    #Add To Archive    ${allfiles}    allpdfs.zip
    Archive Folder With Zip    ${download_Dir}    allpdf.zip
