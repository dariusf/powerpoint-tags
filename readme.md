

# Build

- New presentation
- Tools > Macro > Visual Basic Editor
- Create new module, paste code.vba
- Save as test.pptm
- `./build.sh`, this produces test1.pptm
- Open again and Export > ppam
- Users can now do Tools > PowerPoint Add-ins and select the file

# Enabling macros on macOS

This is required for buttons to work. Manually running macros works out of the box.

- PowerPoint > Preferences > Security > Macro Security > Enable all macros
- Restart

<!--

customui
https://answers.microsoft.com/en-us/msoffice/forum/all/xml-ribbon-powerpoint/5168a260-0941-4260-903a-b783d6782361

docs
https://learn.microsoft.com/en-us/openspecs/office_standards/ms-customui/846e8fb6-07d3-460b-816b-bcfae841c95b

usage
https://www.anirdesh.com/ribbon/callbacks-1.php

mac security problems
https://stackoverflow.com/questions/57733141/developing-a-custom-ribbon-add-in-for-powerpoint-office-365-mac

-->
