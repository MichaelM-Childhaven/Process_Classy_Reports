# Process_Classy_Reports
Takes raw Classy payout report and transaction report and combines them and formats the result

This software is written in VBA for Excel. You can save this macro with a macro-enabled Excel file, but it is recommended that you save it into your personal macro file, which is always available (PERSONAL.XLSB). For more information about creating your PERSONAL.XLSB file and importing macros, please consult the web. The VBA code is in a file called DEVO_MACROS.bas, which you can import into your personal macro file.

If you've never used Excel macro file before, you'll need to learn how to (1) enable the Developer tab in Excel, and then (2) use that Developer tab to import the DEVO_MACROS.bas file. Once you've got everything set up, you can invoke this macro like any other, by pressing Alt-F8 and choosing the correct macro name. The name of the macro is PROCESS_CLASSY_REPORTS.

The idea behind the macro is simply one of efficiency. Classy is an excellent platform for non-profits that provides a wealth of donor data. When a donor makes a gift (contribution), a transaction is made which can then be export in a **Transaction Report**. Secondly, groups of transactions are paid out in **Payout Reports**. The **PROCESS_CLASSY_REPORTS** macro will prompt you to select one Payout Report and one Transaction Report. The idea of course is that the Transaction Report contains transactions that are contained in your Payout Report. This normally entails that the date ranges for your Payout Report and your **Transaction Report (which I called a "Details" Report** are roughly the same, with the Details Report usually having a slightly earlier date range than the Payout Report, because transactions are typically paid out 1-3 days after the transaction occurs.

In Classy, all **Payout Reports** have fixed formats, meaning that you cannot change them. You simply specify the TYPE of the Payout Report (Stripe or PayPal, at the moment) and then give a date range, and that's it. You have no control over the particular data fields that are included, or their order.

The Details Report (the Transaction Report) on the other hand is customizable by you. There is a section in Classy called My Reports. There, you can not only choose the type of transactions (Stripe or PayPal) but you can also specify exactly which columns (fields) you want, and in what order. My PROCESS_CLASSY_REPORTS macro requires a specific set of columns in a specific order, as follows:

1. Transaction Date
