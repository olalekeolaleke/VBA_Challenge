# VBA_Challenge
This Script was used to summarized the stock datasets for the year 2018, 2019 and 2020 respectively.
The first step as seen on the script was creating a "Sub" for the script afterwhich, variables were created.
Titles were created in order to make the Summary Table comprehensible.
A loop was created which was designed to loop through the entire Worksheets and since we are dealing with large dataset -
I decided to use "lastrow" as against being particular about the number of the last row.
The Summary Table consist of four titles which is; the Ticker, the Yearly Change, the Percent Change and the Total Volume.
The Yearly Change and Percent Change values were calculated using simple formulars;
Yearly Change was calculated by deducting the Opening Price from the Closing Price, while the Percent Change was calculated by using (Closing Price/Opening Price - 1)*100
A code was tehn generated to format each cells of the Percent Change based on their values.
If the value equals to or greater than "0", the conditional formatting wa set to "Green", and any cell with a value less than "0" was set to "Red"
