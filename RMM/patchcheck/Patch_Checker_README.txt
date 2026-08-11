Missing Patch Status Checker
============================

1. Install the only dependency once:

   py -m pip install openpyxl

2. Double-click "Start Patch Checker.bat" or run:

   py patch_status_checker.py

3. Select:
   - the older missing-patches report
   - the current patch report
   - the output Excel file

4. Click "Check patches".

Matching method
---------------
The checker matches Client + Device + Patch. Site is deliberately not used because
site names can change or be misspelled between exports.

Results
-------
NOW INSTALLED                  Exact patch is Installed in the current report.
INSTALLED - REBOOT REQUIRED   Exact patch is installed but needs a reboot.
STILL OUTSTANDING              Exact patch is still Missing, Pending, Installing,
                               Failed, or Ignored.
NOT FOUND IN CURRENT REPORT   Exact patch is absent from the current export. It is
                               not automatically treated as installed.

The older report is never overwritten. The output keeps its original columns and
adds Current Patch Status, Current Install Date, and Check Result.
