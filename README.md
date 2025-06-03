# InjectionSpreadsheet
A nice Spreadsheet to track your Estradiol injections :3

Note that this was just designed with personal use in mind, so some things might not be optimized for your exact situation, and frankly that isnt the point of this. Its also only tested with LibreOffice, so that is the only program where I can guarantee anything probably working at all. You may however feel free to just take this and modify it to best fulfill your needs :)

Most important features are the tracking of both injections and used vial with macros for automatically filling out new entries for each with just a dialogue box, for proper utilization of the sheet you will need a precision scale and inform yourself about the density of the contents of your vial.

<h2>Before running code a random stranger let you download (me), please verify that you actually want to run it. While I trust my own code, you should not.</h2>

To use any macros you will need to adjust settings in LibreOffice. Also note that you should never input units, cells are formatted in a way that will automatically add them. Adding them yourself WILL break the sheet.

The Macro "AutoFillNextEmptyRow" will prompt you with an input box for:

- injection number
- date
- time
- vial code
- vial mass in mg before drawing
- vial mass after drawing
- drawn amount in ml
- syringe+needle mass dry
- syringe+needle mass drawn
- syringe+needle mass after injecting

It will then automatically fill the next row, including functions automatically calculating drawn/injected/wasted volume based on density assigned to the vial code.

The Macro "AddNewVial" will prompt you with an input box for:
- vial Code (this has to match with the one you use for injections for that vial, I recommend #1, #2 etc.)
- estimated density (depends on your source, best bet ist to ask them, otherwise make an estimate based on ingredients)
- total amount in ml
- first injection date
- concentration of API in mg/ml
- API
- source of the vial

Pressing "OK" will close the box and automatically add a new vial with all the added info. If you are using a scale, it will also track the total used amount from that vial, so you always know how much you got left.
