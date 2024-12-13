# QuickTranslate
# Copyright Rob Latour 2024
# License MIT

This QuickTranslate program updates the necessary resource files to provide for language support in A Ruler for Windows.

I wrote it myself as an alternative to other more feature rich (but costly) solutions.

It relies on an input "\Location Working Files\Translation.xls" spreadsheet to translate the English resources in A Ruler for Windows into the various languages supported by the program.

QuickTranslate uses the English (en-US) values in the "\Location Working Files\Translation.xls" spreadsheet (column A) in both the "InCode" and "InForm" worktabs, and the English (en-US) resource files in A Ruler for Windows as its base.

If within A Ruler for Windows you want to make a change to any language variable you should:

- make it in the A Ruler for Windows code / form design 
- rebuild A Ruler for Windows
- make the change in the "\Location Working Files\Translation.xls" spreadsheet
- run the QuickTranslate program, and 
- rebuild A Ruler for Windows

The .xls is also released as part of the open source A Ruler for Widows project https://github.com/roblatour/ARulerForWindows/ and can be found at: https://github.com/roblatour/ARulerForWindows/blob/main/Localization%20Working%20Files/Translation.xls
