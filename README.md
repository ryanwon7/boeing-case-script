<h1> Boeing Case Competition Python Excel Script </h1>
This project is a small python script written by Ryan Won for the 2019 Drexel University Boeing Case Competition.

<p><h3>Purpose</h3>
One of the provided issues for the Boeing Case Competition is to reduce the cost to create the Chinook aircraft that Boeing produces. To do so, competitiors are provided with a list of suppliers for all the parts of the aircraft; this is listed in the repository files as "suppliers.xslx" In this spreadsheet, three different supplier options all with differing price per part, lead time, quality acceptance %, and on-time delivery % are listed for the 20+ general parts that the Chinook is composed of. This script is written to analyze each supplier and their respective factors and determine which supplier Boeing should order from in order to assemble the most cost effective aircraft.

<p><h3> Mathematical Formula for "Impact Cost"</h3>
In order to compare each supplier and determine which of the options is the most cost effective, I have developed a formula to calculate what I call the "impact cost" of each part. The formula for the impact cost is listed below:
<h4> Impact Cost = QTY * CPP * (1 + ((1 - OTD%) * LT)/4 +  (1 - QA%)) </h4>
Where
<p> QTY = number of the specific part required to the assembly of one Chinook
<p> CPP = the cost in dollars for one part
<p> OTD% = the on-time delivery percentage as a decimal (i.e. 98% => 0.98)
<p> LT = lead time in months
<p> QA% = the quality acceptance percentage as a decimal (i.e. 99% => 0.99)
  
<p> <h3> Formula Reasoning </h3>
