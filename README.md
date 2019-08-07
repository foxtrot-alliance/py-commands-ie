# py-commands-ie

<i>Still under initial development...</i>

For support, please contact us: support@foxtrotalliance.com

This program allows you to execute commands in Internet Explorer such as clicking, getting values, setting values, sending values, and more. In some cases, it is not possible to effectively engage with elements on a website in Internet Explorer with the standard targeting technology of Foxtrot. Or, in some cases you simply aim to increase speed or precision. In such scenarios, this program comes in handy as you can create custom logic to perform actions in websites. You can run the program via the CMD or as part of an automation script in an RPA tool like Foxtrot. This solution is meant to supplement Foxtrots core email functionality and enable you to find elements very fast and perform commands not possible using the standard Foxtrot browser technology. The solution is written in Python using the modules "pywin32", "pyautogui", and "keyboard". You can see the [full source code here](https://github.com/foxtrot-alliance/py-commands-ie/blob/master/py-commands-ie.py).

## Installation

1. Download the [latest version](https://github.com/foxtrot-alliance/py-commands-ie/releases/download/v0.0.5/py-commands-ie_v0.0.5.zip).
2. Unzip the folder somewhere appropriate, we suggest directly on the C: drive for easier access. So, your path would be similar to "C:\py-commands-ie_v0.0.5".
3. After unzipping the files, you are now ready to use the program. The only file you will have to be concerned about is the actual .exe file in the folder, however, all the other files are required for the solution to run properly.
4. Open Foxtrot (or any other RPA tool) to set up your action. In Foxtrot, you can utilize the functionality of the program via the DOS Command action (alternatively, the Powershell action).

## Usage

When using the program via Foxtrot, the CMD, or any other RPA tool, you need to reference the path to the program exe file. If you placed the program directly on your C: drive as recommended, the path to your program will be similar to: 
```
C:\py-commands-ie_v0.0.5\py-commands-ie_v0.0.5.exe
```
TIP: Make sure NOT to surround the path with quotation marks in your commands.

## Commands

<i>More information coming...</i>

<i>Temp examples:</i>
```
http://ec.europa.eu/taxation_customs/vies/:
EXE_PATH -find_element1 "id=countryCombobox" -command "select" -value "DK"
EXE_PATH -find_element1 "id=number" -command "send" -value "1234556678954645"
EXE_PATH -find_element1 "id=submit" -command "click"
EXE_PATH -find_element1 "id=submit" -command "click_bypass"

http://www.nationalbanken.dk/da/statistik/valutakurs/Sider/default.aspx:
EXE_PATH -find_element1 "id=currenciesTable" -find_element2 "tagname=tr" -command "count"
EXE_PATH -find_element1 "id=currenciesTable, item=1" -find_element2 "tagname=tr" -find_element3 "tagname=td" -command "get" -attribute "innerText"
EXE_PATH -find_element1 "id=currenciesTable" -find_element2 "classname=icons-xls" -find_element3 "parent=true" -command "get" -attribute "href"
```
