# RTD-Extension

Version 1.0.0 is here

## Overview

This is a Python based extension for LibreOffice Calc to make real-time data available in Calc spreadsheets.
It is a simple implementation similar to Microsoft Excel RTD() function

This implementation is based on https://github.com/cmallwitz/Financials-Extension
Cheers mate, it would've been impossible without your code

### Feedback requested:

Please provide feedback about using the extension [here](https://github.com/gilgil/RTD-Extension/issues)

### Usage:

Under 'Releases' on GitHub [there](https://github.com/gilgil/RTD-Extension/releases) is a downloadable **RTD-Extension.oxt** file - load it into Calc under menu item: Tools, Extension Manager...

Please make sure, not to rename the OXT file when downloading and before installing: LO will mess up the installation otherwise and the extension won't work.

1. Getting data should be as simple as having this in a cell: `=RTD("key")`

2. Now, in a different cell, type `=RTD("__start")` and a TCP listener will be started waiting for incoming data.

3. Next, data needs to be sent using a socket to port 13000.
   If you wish to manually test it, open up a telnet in a terminal: `telnet localhost 13000`.
   In the telnet, write '<key|example_data>' and then press '<Enter>'.
   The cell with the formula `=RTD("key")` (from section 1) will now show "example_data" instead of "None"

### Build:

You will need the LibreOffice SDK installed. 

On my system (Ubuntu) I downloaded both LibreOffice & LibreOffice SDK from https://www.libreoffice.org/ and followed the instructions

\# depending on your location...

cd ~/RTD-Extension/

\# This builds file **RTD-Extension.oxt**

./compile.sh

### Tested with:
- Ubuntu 18.04.5 / LibreOffice Calc 7.4.2.3 / Python 3.8.14
