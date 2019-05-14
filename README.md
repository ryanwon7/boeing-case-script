<h1> Boeing Case Competition Python Excel Script </h1>
This project is a small python script written by Ryan Won for the 2019 Drexel University Boeing Case Competition.

<p><h3>Installation and Run</h3><p>
This script was written and tested in Python 2.7.15. Compatability for older versions of Python 2.7 or other versions of Python untested.

This script requires the python os and openpyxl modules. The openpyxl module requires the pip package manager to be installed, and pip should come shipped with Python 2.7 on versons 2.7.9+.

To run in command line run "python excel-script.py", assuming python is the identifier to run Python 2.7.

<p><h3>Purpose</h3>
One of the provided issues for the Boeing Case Competition is to reduce the overall cost of building the Chinook aircraft that Boeing produces. To do so, competitiors are provided with a list of suppliers for all the parts of the aircraft; this is listed in the repository files as "suppliers.xslx". In this spreadsheet, three different supplier options, all with differing price per part, lead time, quality acceptance %, and on-time delivery %, are listed for the 20+ general parts that the Chinook is composed of. This script is written to analyze each supplier and their respective factors and determine which supplier Boeing should order from in order to assemble the most cost effective aircraft.

<p><h3> Mathematical Formula for "Impact Cost"</h3>
In order to compare each supplier and determine which of the options is the most cost effective, I have developed a formula to calculate what I call the "impact cost" of each part. The most stringent formula for the impact cost is listed below:
<h4> Impact Cost = QTY * CPP * (1 + ((1 - OTD%) * LT) +  (1 - QA%)) </h4>
Where
<p> QTY = number of the specific part required to the assembly of one Chinook
<p> CPP = the cost in dollars for one part
<p> OTD% = the on-time delivery percentage as a decimal (i.e. 98% => 0.98)
<p> LT = lead time in months
<p> QA% = the quality acceptance percentage as a decimal (i.e. 99% => 0.99)
  
<p> <h3> Formula Reasoning </h3>
<p> What this formula does is it takes the base price, QTY * CPP, and then adds extra cost based on how far away the OTD% and QA% are from 100%, and on how long the lead time is. 
<p> The reasoning for this is that a low OTD% will lead to more late deliveries, which can throw the tight schedule of plane production off balance and incur late fees when doing completed Chinook deliveries (which in the competition parameters, is 5% off the sale price for each month late). A similar reasoning goes for the QA%, as a lower quality acceptance rating will lead to more rejected parts which in turn takes time to fix or get a new part, delaying the process. This again leads to costruction process delays and possible late fees. Lead Time is a factor as it affects how early in the process a part needs to be ordered, and also is a factor in how quickly a part can come when it is late or fails QA. 

<p> <h3> Formula Adjustments </h3>
<p> The formula stated above is the most stringent as it places the most weight on the OTD%, LT, and QA%. If one would want to adjust the formula to put less weight on one of those factors and more weight on the original price, you can divide the respective factor by a numerical factor. For example, an output that focuses on price over quality would look something like this: 
<h4> Impact Cost = QTY * CPP * (1 + ((1 - OTD%) * LT)/4 +  (1 - QA%))/4 </h4>
<p> Thus, the formula can be adjusted by the user in the script to achieve the desired analysis and output.
