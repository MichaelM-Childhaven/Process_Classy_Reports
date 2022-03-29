# Process_Classy_Reports v1.0
**SUMMARY: Takes two files -- a raw Classy Payouts Report export (a CSV file) and a Transactions Report export (a CSV file) -- and combines them and formats the result into one useful, well formatted, time-saving file that you can then save as a normal Excel file. And the data is ready to be copy-pasted into Raiser's Edge or another fundraising database of your choice.**

This software is written in VBA for Excel. You could save this macro inside a macro-enabled Excel file, but it is recommended that you save it into your personal macro file, which is always available (PERSONAL.XLSB). For more information about creating your PERSONAL.XLSB file and importing macros, please consult the web. The VBA code is in a file called **DEVO_MACROS.bas**, which you can import into your personal macro file (your PERSONAL.XLSB file).

If you've never used Excel macros file before, you'll need to learn how to (1) enable the Developer tab in Excel, and then (2) use that Developer tab to import the DEVO_MACROS.bas file. Once you've done that, you can invoke this macro like any other, by pressing Alt-F8 and choosing the correct macro name. The name of the macro is **PROCESS_CLASSY_REPORTS**.

The idea behind the macro is simply one of efficiency. Classy is an excellent platform for non-profits that provides a wealth of donor data. When a donor makes a gift (contribution), a transaction is recorded which can then be exported in a **Transaction Report** from Classy (what I call a **"Details Report"**). Secondly, groups of transactions are paid out in **Payout Reports**. The **PROCESS_CLASSY_REPORTS** macro will prompt you to select one Payout Report and one Details Report. The idea of course is that the Details Report contains transactions that are contained in your Payout Report. This normally entails that the date ranges for your Payout Report and your Details Report are roughly the same, with the Details Report usually having a slightly earlier date range than the Payout Report, because transactions are typically paid out 1-3 days after the transaction occurs.

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

So when you create your report, you want to specify those columns in that order. You may want have two such Details Reports, one for Stripe transactions and one for PayPal transactions (if your organization accepts both payment methods), because these two types of transactions are paid separately in their own separate Payout Reports.

You'll also need to set a filter for your Details Report. To specify a Stripe Details Report, in your My Report setup, set a **Filter** with **Property = Payment Processor**, **Operator = Is Equal To** and **Value = 1 Value = Classy Pay Powered by Stripe** OR **Value = 1 Value = PayPal Commerce**. 

Classy will let you then set a date range and download (export) your Details Report immediately; for the Payout Report, you set a date range and then export the report, and then Classy sends you an e-mail containing a link to download it (which normally comes within a couple of minutes).

Once you have your Payout and Details reports downloaded, running the macro is simple: Run the PROCESS_CLASSY_REPORTS macro, select the payout and details reports, and it does the rest. When it finishes, it will give you a message that you should now save the resulting file as an Excel file to preserve the formatting. (Your exported Classy Payout and Details reports are in CSV format.)

Here's what the macro does:

1. It combines the data fields that are contained only in the Details report into the Payout report, creating one report that has everything you need for gift entry into Raiser's Edge or another fundraising database platform. It does this by matching Details Report transactions with Payout Report transactions via the Payment Processor Reference ID, which will be the same in both reports.

2. It nicely formats all of the data, including the dreaded "East Coast zipcodes" that begin with one or more leading zeroes. Note that the leading-zero problem is not a Classy problem, but rather an Excel problem: Excel looks at your CSV file and takes anything that **looks** like a number with leading zeroes and immediately removes those leading zeroes. My macro KNOWS that this should be a zipcode and converts leading-zero zipcodes to text fields with their leading zeroes preserved. The macro code works well with both USA and Canadian zip codes.
 
3. It combines the **Billing Address** and **Billing Address 2** fields into just the Billing Address field. For most Raiser's Edge users, the "street address" portion is just a single line and a single field, so this is much more convenient for RE.

4. It takes the TEXTUAL fields **Dedication Type, Dedication Name, Dedication Message and Donor's Comment** and cleans up the text and combines it into a single handy field that you can simply cut and paste into your RE Reference field. Not only that, but it also adds text that indicates whether it's a Stripe or PayPal transaction, and whether it's a one-time gift or a recurring gift. All in one convenient field that you can simply cut and paste into your RE Reference field.

5. It sorts the transactions by Payout Date firstly and by Transaction Date secondly, and then groups the sorted transactions by Payout Date. The resulting groups are then summed, automatically giving you totals for both the **Gross Transaction Amount** and the **Net Amount**. Note that sometimes Classy will make *more* than one payout on a single day, and my macro would have no way of knowing about that, so sometimes you may need to further divvy up these groups. What's nice is, these totals are _formulas_, not hard coded, so you can copy and repurpose them when you need to split up a group.

7. It formats your amounts to look like amounts (with currency formatting) and your phone numbers to look like phone numbers.

8. It adjust column widths and in general tries to lay everything out nicely for you.

As mentioned earlier, the macro does the work of matching up detail transactions row to payout transaction rows. The payout report actually contains most of the data we want already, with the following exceptions: **Dedication Type, Donor Is Anonymous, Donor Phone Number, and Donor's Comment**. When the PROCESS_CLASSY_REPORTS macro finishes, it creates a file with a header row colored yellow _except for some orange-colored fields at the end_. Those orange-colored fields represent the data that was taken from the Details Report (from one of those four columns I just mentioned) and combined into this "new, enhanced" payout report. That means that **_if there was a match, the orange Reference field (and possibly other orange fields) will contain some data._** If the orange reference field for a given payout transaction contains no data, that means that no match for that transaction was found in your Details report. If on the other hand there were transactions in the Details report that didn't match any rows in the Payout report, then those transactions were simply ignored.

# Bonus Macros

There are two other small macros included in the DEVO_MACROS.bas file for your review and use. For more info about them, see the description field for the DEVO_MACROS.bas file. Basically there's one that just does a quick format of a header row in most any downloaded CSV file, and the other one "preps" your formatted Excel file to be saved as a plain CSV file, by deleting extraneous rows and columsn (so you don't end up with extra junk fields in your CSV file). Again, see the DEVO_MACROS.bas file description for more info.

I truly hope that other non-profits find this tool useful! If you have any questions at all, please do not hesitate to ask me.

<div>Michael Matloff<br>
Development Systems Manager<br>
Childhaven<br>
(Childhaven.org)</div>
