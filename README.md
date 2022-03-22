# Process_Classy_Reports
Takes raw Classy payout report and transaction report and combines them and formats the result

This software is written in VBA for Excel. You can save this macro with a macro-enabled Excel file, but it is recommended that you save it into your personal macro file, which is always available (PERSONAL.XLSB). For more information about creating your PERSONAL.XLSB file and importing macros, please consult the web. The VBA code is in a file called DEVO_MACROS.bas, which you can import into your personal macro file.

If you've never used Excel macro file before, you'll need to learn how to (1) enable the Developer tab in Excel, and then (2) use that Developer tab to import the DEVO_MACROS.bas file. Once you've got everything set up, you can invoke this macro like any other, by pressing Alt-F8 and choosing the correct macro name. The name of the macro is PROCESS_CLASSY_REPORTS.

The idea behind the macro is simply one of efficiency. Classy is an excellent platform for non-profits that provides a wealth of donor data. When a donor makes a gift (contribution), a transaction is made which can then be export in a **Transaction Report**. Secondly, groups of transactions are paid out in **Payout Reports**. The **PROCESS_CLASSY_REPORTS** macro will prompt you to select one Payout Report and one Transaction Report. The idea of course is that the Transaction Report contains transactions that are contained in your Payout Report. This normally entails that the date ranges for your Payout Report and your **Transaction Report (which I called a "Details" Report** are roughly the same, with the Details Report usually having a slightly earlier date range than the Payout Report, because transactions are typically paid out 1-3 days after the transaction occurs.

In Classy, all **Payout Reports** have fixed formats, meaning that you cannot change them. You simply specify the TYPE of the Payout Report (Stripe or PayPal, at the moment) and then give a date range, and that's it. You have no control over the particular data fields that are included, or their order.

The Details Report (the Transaction Report) on the other hand is customizable by you. There is a section in Classy called My Reports. There, you can not only choose the type of transactions (Stripe or PayPal) but you can also specify exactly which columns (fields) you want, and in what order. My PROCESS_CLASSY_REPORTS macro requires a specific set of columns in a specific order, as follows:

1. Transaction Date
2. Supporter Name
3. Transaction Status
4. Frequency
5. Gross Transaction Amount
6. Net Transaction Amount
7. Dedication Type
8. Dedication Name
9. Dedication Message
10. Donor is Anonymous
11. Donor Phone Number
12. Donor's Comment
13. Payment Processor
14. Payment Processor Reference ID

So when you create your report, you want to specify those columns in that order. You'll want to have at least two such Details Reports, one for Stripe transactions and one for PayPal transactions, because these two types of transactions are paid separately in their own separate Payout Reports.

So once you invoke the macro, you'll specify a Payout Report (of either the Stripe or PayPal variety) and a Details Report (produced from your My Reports report, and again of either the Stripe or PayPal variety). To specify a Stripe Details Report, in your My Report setup, set a **Filter** with **Property = Payment Processor**, **Operator = Is Equal To** and **Value = 1 Value = Classy Pay Powered by Stripe** OR **Value = 1 Value = PayPal Commerce**. 

Classy will let you then set a date range and download (export) your Details Report immediately; for the Payout Report, you set a date range and then export the report, and Classy notifies you via e-mail containing a link to download it (which is normally not more than a couple minutes later).

Once you have your Payout and Details reports downloaded, running the macro is simple: Run the PROCESS_CLASSY_REPORTS macro, select the payout and details reports, and it does the rest. When it finishes, it will give you a message that you should now save the resulting file as an Excel file to preserve the formatting. (Your exported Classy Payout and Details reports are in CSV format.)

Here's what the macro does:

1. It combines the data fields that are contained only in the Details report into the Payout report, creating one report that has everything you need for gift entry into Raiser's Edge or another fundraising database platform. It does this by matching Details Report transactions with Payout Report transactions via the Payment Processor Reference ID, which will be the same in both reports.

2. It nicely formats all of the data, including the dreaded "East Coast zipcodes" that begin with one or more leading zeroes. Note that the leading-zero problem is not a Classy problem, but rather an Excel problem: Excel looks at your CSV file and takes anything that **looks** like a number with leading zeroes and immediately removes those leading zeroes. My macro KNOWS that this should be a zipcode and converts leading-zero zipcodes to text fields with their leading zeroes preserved. The macro code works well with both USA and Canadian zip codes.
 
3. It combines the Address 1 and Address 2 fields into just the Address 1 field. For most Raiser's Edge users, the address is just a single line and a single field, so this is much more convenient for RE.

4. It takes all of the TEXTUAL fields **Dedication Type, Dedication Name, Dedication Message and Donor's Comment** and cleans up the text and combines it into a single handy field that you can simply cut and paste into your RE Reference field. Voila!

5. It sorts the transactions by Payout Date firstly and by Transaction Date secondly, and then groups the sorted transactions by Payout Date. The resulting groups are then summed, automatically giving you totals for both the **Gross Transaction Amount** and the **Net Amount**. Note that sometimes Classy will make *more* than one payout on a single day, and my macro would have no way of knowing about that, so sometimes you may need to further divvy up these groups. What's nice is, these totals are _formulas_, not hard coded, so you can copy and repurpose them when you need to split up a group.

7. It formats your amounts to look like amounts (with currency formatting) and your phone numbers to look like phone numbers.

8. It adjust column widths and in general tries to lay everything out nicely for you.

I hope that other non-profits find this tool useful!

Michael Matloff
Development Systems Manager
Childhaven
(Childhaven.org)
