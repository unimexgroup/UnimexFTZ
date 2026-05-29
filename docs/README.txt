UNIMEX FTZ PROCESSOR
====================

WHAT IT DOES
------------
Takes a shipment master manifest and its FTZ separation list, and produces
a condensed FTZ Excel file sorted by HS code with summed quantity, weight,
and value.

HOW TO USE
----------
1. Put UnimexFTZ.exe in any folder you like.

2. Double-click UnimexFTZ.exe once. It will create three folders next to it:
       input\    - drop your shipment files here
       output\   - finished FTZ files appear here
       logs\     - run logs (for troubleshooting)

3. Drop BOTH files for each shipment into the "input" folder:
       - the master manifest (the big file with all the line items)
       - the separation list (the small file with bag IDs and tracking #s)

   You can put multiple shipments in at once. The tool figures out which
   files go together based on the shipment ID number inside each file.

4. Double-click UnimexFTZ.exe again. It will:
       - read every file in input\
       - pair up masters with their separation lists
       - produce one {ShipmentID}_FTZ.xlsx file per shipment in output\

5. When done, the window shows a summary and waits for you to press ENTER.

WHAT THE OUTPUT LOOKS LIKE
--------------------------
Each output file has 6 columns:
   HS Code   |   Quantity   |   Weight   |   Value   |   Zone   |   Charges

- One row per unique HS code (line items are summed within each code)
- Sorted ascending by HS code
- Zone is always "P", Charges is always 3
- Any final Value < $1 or Weight < 1 kg is rounded up to 1

IF SOMETHING LOOKS WRONG
------------------------
- Check the logs\ folder for the most recent run log
- Send that log file (and the inputs you used) to Andy

WHAT THE TOOL WILL TELL YOU
---------------------------
- [OK   ] = shipment processed successfully
- [SKIP ] = couldn't process this shipment (e.g. master found but no
            separation list, or vice versa)
- [WARN ] = something to be aware of (e.g. a bag ID in the separation
            list wasn't found in the master)
- [?????] = a file in the input folder wasn't recognized as either a
            master or a separation list

QUESTIONS / ISSUES
------------------
Contact Andy at Unimex.
