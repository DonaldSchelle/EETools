# **Automate those Tedious Resistor Calculations**

By Don Schelle - January 2021



Once again, I find myself travelling down a path many of you have surely walked before.  Inspiration strikes and I embark on a journey to create the next great circuit design.  After hours of research, I start calculating some component values and fire up Excel to automate the more complex computations.   When scanning the results, an all too familiar sting hits my posterior; a calculated resistor value is 123.4567 kΩ.    A trivial problem as I pull out my favourite set of E96 tables<sup>1</sup> and select the closest value (123 kΩ).   After recalculating the spreadsheet with the standard value, I tinker a little further, resulting in a new calculated value of 765.4321 kΩ.   Back to the tables, rinse, repeat, ad nauseam.   There has to be a better way.

Candidly, this problem has been solved many different ways with cumbersome large equations<sup>2</sup>, and VBA excel add-ins<sup>3,4</sup>.   The former requires manually entering large equations, while the latter requires that you not only have the Excel add-in, but also remember how to use it, and turn on macro-enabled spreadsheets.   An excel tool (ExcelDNA<sup>5</sup>) provides a framework to create a simple yet flexible user-defined-function (UDF) that seamlessly integrates into Excel via a .XLL add-in.   Pivotal is that ExcelDNA enables the developer to support IntelliSense (figure 1), which adds interactive help statements, easing UDF usage, and making it operate more like a built-in Excel function.   Let’s get started!



<img src="Figures\Figure 1.png" style="zoom: 50%;" />

<p align="center"><b>Figure 1</b> – ExcelDNA enables support for IntelliSense which provides a handy interactive reference when using Excel UDFs</p>





EETtools is a .XLL that leverages the ExcelDNA framework to create three UDF’s that ease resistor calculations.   The simplest of the three functions, ***ESERIES***, selects an appropriate value from a standard E-series value table.   After choosing a target value (dValue), set the sSeries parameter to select which table is used (E3 to E192).  A speed-optimized binary search through the E-series table selects the next lowest and highest E-series value.   Offset error for each selected value is calculated, and the function returns the E-series value yielding the lowest calculated error.  Table 1 outlines all available *ESERIES* UDF parameters.



**Table 1. *ESERIES* UDF Parameters**

|**Parameter**|**Default**|**Description**|
| :-: | :-: | :- |
|dValue||Target value to convert|
|sSeries||Desired E-Series Table|
|[*iRound*]|0|Rounding Option <br/>-1 = Choose the closest lower or equal value <br/> 0 = Choose the closest value (higher, lower, or equal)<br/> 1 = Choose the closest higher or equal value|
|[*iCalcType*]|0|Offset Calculation Method<br/>0 = Calculate the closest algebraic value<br/>1 = Calculate the closest percent error value<br/>2 = Calculate the closest percent difference value<br/>3 = Calculate the closest logarithmic value|
|[*dMinLimit*]|1|Limit the returned value to no less than dMinLimit|
|[*dMaxLimit*]|1,000,000|Limit the returned value to no more than dMaxLimit|

- [  ]  - Parameters are *optional*



### **Calculating Error**

All functions in the EETools library seek to minimize the amount of error when calculating target component values.   Optional iRound and iCalcType parameters determine how error is calculated.  Careful selection of these parameters can tailor the returned results to a desired application.  

Use the iRound parameter when the desired outcome must be limited to a ceiling or floor value.  Force the result to either the closest value (iRound = 0, *default*), the closest equal or lower value (iRound = -1), or the closest higher or equal value (iRound = +1).   In other words, iRound forces the function to consider only negative, only positive, or both positive and negative error.

When iRound is set to 0, all functions consider both positive and negative error.   Setting the iCalcType (table 2) parameter selects the desired equation that calculates offset error.  In the case of logarithmic error, a simplification of the error equation optimizes computational performance.



**Table 2. iCalcType Summary**

| **iCalcType** | **Description**    | **Calculation**                                                 |
| :-----------: | :-----------------:| :-- |
|       0       | Algebraic          | <div align="center"><img style="background: white;" src="https://render.githubusercontent.com/render/math?math=\delta = \lvert \nu_A-\nu_E \rvert"></div> |
|       1       | Percent Error      | <div align="center"><img style="background: white;" src="https://render.githubusercontent.com/render/math?math=%5Cdelta%20%3D%20%5CBigg%5Clvert%7B%5Cfrac%7B%5Cnu_A-%5Cnu_E%7D%7B%5Cnu_E%7D%5CBigg%5Crvert%7D%5Ctimes%20100%5C%25"></div>|
|       2       | Percent Difference | <div align="center"><img style="background: white;" src="https://render.githubusercontent.com/render/math?math=%5Cdelta%20%3D%20%5CBigg%5Clvert%7B%5Cfrac%7B%5Cnu_A-%5Cnu_E%7D%7B%5Cfrac%7B(%5Cnu_A%2B%5Cnu_E)%7D%7B2%7D%7D%5CBigg%5Crvert%7D%5Ctimes%20100%5C%25"></div>|
|       3       | Logarithmic Error  | <div align="center"><img style="background: white;" src="https://render.githubusercontent.com/render/math?math=%5Cdelta%20%3D%20%5CBigg%5Clvert%7B%5Cfrac%7Blog(%5Cnu_A)-log(%5Cnu_E)%7D%7B2%7D%5CBigg%5Crvert%7D%20%3D%20%5CBigg%5Clvert%5Cfrac%7B1%7D%7B2%7D%5Ctimes%20log%5Cleft(%5Cfrac%7B%5Cnu_A%7D%7B%5Cnu_E%7D%5Cright)%20%5CBigg%5Crvert"></div> |


where

- &delta; is calculated error
- &nu;<sub>A</sub> is actual value (test value)
- &nu;<sub>E</sub> is expected value (ideal value)



Algebraic and Percent Error yield equally weighted error regardless of positive or negative error.  Use Percent Difference and Logarithmic Error for applications that benefit by favouring positive error over negative.   Plotting error calculation results (figure 2) using an arbitrary 1Ω resistor yields additional insight.  With a large enough range, the non-linear behavior of the Percent Difference and Logarithmic calculation methods become clear. 



<img src="Figures\Figure 2.png" style="zoom:50%;" />

<p align="center"><b>Figure 2</b> – EETools allows the user to choose how error is calculated and can favour positive over negative error.  Percent and Percent Difference error are plotted on the left axis.   Algebraic and Logarithmic error are plotted on the right axis.</p>



Further illustrating; consider the value of 1009.96 Ω using an E96 table (figure 3).   Algebraically the value is closer to 1000Ω, though logarithmically the value is closer to 1020Ω.  When using practical values, the difference between calculation types is subtle, with the ideal option varying according to application requirements.  In most cases, a simple algebraic function is sufficient and, due to simplicity/speed, preferred if there are multiple calculations to make on a given spreadsheet.

<img src="Figures\Figure 3.png" style="zoom: 50%;" />

<p align="center"><b>Figure 3</b> – Application requirements may require one of many supported offset calculation types.</p>



### **Resistors Pairs**

There are times when a precision resistor value is required but is not available as a standard value.   Traditional wisdom would dictate choosing an arbitrary resistor value (i.e. 10 kΩ) in series with a second smaller value to close the gap.  

*ESeriesResistorPair* UDF automatically calculates ideal values while guaranteeing the closest match possible by searching through all available combinations in the chosen E-series table according to the iRound and iCalcType parameters.  Set the iComboType parameter to search for either a series (iComboType = 0) or parallel (iComboType = 1) pair resistor network.



Generally, Excel UDF’s only return a single variable and this function is no different.   Setting iReturnValue selects the returned component value.   Though considerable time was spent optimizing all UDF’s for speed, returning a single value requires the function to be executed twice, thus increasing processing time.   To maximize speed, instruct *ESeriesResistorPair* UDF to return an array of values by setting iReturnValue to 0.  Implement the array feature by first selecting a number of cells in either a horizontal row, or vertical column.  Enter the function in the formula bar and press **CTRL**+**SHIFT**+**ENTER** to create the final array of output values (figure 4).   



<img src="Figures\Figure 4.png" style="zoom:50%;" />

<p align="center"><b>Figure 4</b> – Excel returns an array of values when pressing <b>CTRL</b>+<b>SHIFT</b>+<b>ENTER</b> after selecting the desired cell configuration, and entering a supported function in the formula bar.</p>



Sometimes multiple resistor combinations can achieve the same result, perfectly matching the required target value.  *ESeriesResistorPair* implements a secondary search parameter to further fine-tune the output.   Setting parameter dMatchType to -1 performs a comprehensive search and picks resistor values that are closest to each other.    Alternatively, setting dMatchType to a positive value, will select a resistor pair with one of the resistor values closest to the dMatchType value.   This can be handy when trying to optimize a network that has a parallel capacitor on one of the series elements.   Setting dMatchType to 0, optimizes processing speed by exiting after the first perfect match is returned.   

Without physical limits, resistor values may be chosen that simply aren’t practical.    For example, series resistor pairs might use components in the mΩ range, while parallel pairs might use components in the GΩ range; neither of which are ideal.   Use the optional dCompMinLimit and dCompMaxLimit parameters to ensure that all returned component values are within a practical range for final circuit synthesis.  Table 3 outlines all available *ESeriesResistorPair* UDF parameters.



**Table 3. *ESeriesResistorPair* Function Variables**


|**Parameter**|**Default**|**Description**|
| :-: | :-: | :- |
|dValue||Target value|
|sSeries||Desired E-Series Table|
|iComboType||Resistor Network<br/>0 = Series Resistor Pair <br/>1 = Parallel Resistor Pair|
|iReturnValue||Return Value <br/>0 = Return an array of values.  Highlight desired number of cells, enter formula, and press **CTRL**+**SHIFT**+**ENTER**<br/>1 = Return a single value for resistor #1<br/>2 = Return a single value for resistor #2<br/>3 = Return a single value for the algebraic difference between the target value and Thevenin resistance of the calculated values|
|[*iRound*]|0|Rounding Option<br/>-1 = Choose the closest lower or equal value <br/>0 = Choose the closest value (higher, lower, or equal)<br/>1 = Choose the closest higher or equal value|
|[*iCalcType*]|0|Offset Calculation Method<br/>0 = Calculate the closest algebraic  value  <br/>1 = Calculate the closest percent error value<br/>2 = Calculate the closest percent difference value<br/>3 = Calculate the closest logarithmic value<br/>|
|[*dMatchType*]|0|Matching when equal offsets are found<br/>-1 = Choose values that are closest to each other<br/>0 = Choose first available perfect match <br/>> 0 = Choose best pair with one value closest to dMatchType|
|[*dCompMinLimit*]|1|Limit any resistor value to no less than dCompMinLimit|
|[*dCompMaxLimit*]|1,000,000|Limit any resistor value to no more than dCompMaxLimit|

- [  ]  - Parameters are *optional*



### **Resistor Ratios!**

Generating resistor values given a desired target ratio is the next evolution.   Use the *ESeriesResistorRatio* to accomplish exactly this.   

iRatioType configures the UDF to return resistor values tailored to a specific application (figure 5).  Set iRatioType to **Simple Ratio** (iRatioType = 0) instructs the calculator to find resistors that will most closely match the desired ratio.   Setting iRatioType to **Voltage Divider** (iRatioType = 1) instructs the calculator to first convert the voltage ratio into an equivalent resistor ratio, and then search for it.  For example, suppose we want a voltage divider ratio D = R<sub>SECONDARY</sub> / (R<sub>PRIMARY</sub> + R<sub>SECONDARY</sub>), the calculator first converts it to resistor ratio R<sub>SECONDARY</sub>/R<sub>PRIMARY</sub> = D / (1-D), and then searches for that ratio.

<img src="Figures\Figure 5.png" style="zoom:50%;" />

<p align="center"> <b>Figure 5</b> – The <i>ESeriesResistorRatio</i> UDF can be configured to find values for either a simple resistor ratio, typical when calculating gain for an inverting op-amp; or for a voltage divider ratio typical for feedback resistor calculations when implementing power regulators.</p>



*ESeriesResistorRatio* maintains similar functionality with the optional iRound, iCalcType, and dCompMinLimit/dCompMaxLimit.   Also included in the function are two additional parameters.   dThevMinLimit/dThevMaxLimit forces the function to return values that not only meet the desired ratio, but also ensure that the Thevenin resistance of the network is goverened to desired limits.   Use the dMatchType parameter to tune the Thevenin resistance of the entire network.  These parameters are especially useful when the application requires a minimum, maximum, or preferred current flowing through the network.   Setting iElements enables calculation for 5 different network types (figure 6).  Combining series or parallel elements create networks that are much closer to, and many times perfectly match, the target ratio.  When iElements is set to 1 through 4, *ESeriesResistorRatio* leverages the *ESeriesResistorPair* UDF to calculate the best combination.  In these cases, a comprehensive search for *ESeriesResistorPair* values is forced, ensuring the most practical results should multiple series/parallel elements produce the same result.

<img src="Figures\Figure 6.png" style="zoom:50%;" />

<p align="center"><b>Figure 6</b> – The iElements parameter instructs the <i>ESeriesResistorRatio</i> to find values according to the desired network implementation.</p>



As an example (figure 7), we can use the iElements parameter (figure 6) to calculate feedback resistors that generate a 3.3V output with a 0.8V feedback threshold.   A simple resistor pair won’t yield an exact ratio, thus iElements has been set to use a combined series element (iElements = 1) on the bottom side of the divider, yielding the desired perfect ratio.  Table 4 outlines all available *ESeriesResistorRatio* UDF parameters.

<img src="Figures\Figure 7.png" style="zoom:50%;" />

<p align="center"><b>Figure 7</b> – The iElements parameter allows the <i>ESeriesResistorRatio</i> to generate perfect resistor ratio networks under nearly all circumstances.</p>





**Table 4. *EseriesResistorRatio* Function Variables**

|**Parameter**|**Default**|**Description**|
| :-: | :-: | :- |
|dNumerator||Numerator of desired ratio|
|dDenominator||Denominator of desired ratio|
|sSeries||Desired E-Series Table|
|iRatioType||Ratio Type<br/>0 = Simple Ratio<br/>1 = Voltage Divider Ratio|
|iReturnValue||Return Value<br/>0 = Return an array of values.  Highlight desired number of cells, enter  formula, and press **CTRL**+**SHIFT**+**ENTER**<br/>1 = Return a single value for R<sub>PRIMARY</sub><br/>2 = Return a single value for R<sub>SECONDARY</sub><br/>3 = Return a single value for R<sub>TERTIARY</sub><br/>4 = Return a single value error between the target ratio and the calculated E-series ratio, calculated according to iCalcType|
|[*iRound*]|0|Rounding Option<br/>-1 = Generate closest lower or equal ratio<br/>0 = Generate closest ratio (higher, lower, or equal)<br/>1 = Generate closest higher or equal ratio|
|[*iCalcType*]|0|Offset Calculation Method<br/>0 = Calculate the closest algebraic value<br/>1 = Calculate the closest percent error value<br/>2 = Calculate the closest percent difference value<br/>3 = Calculate the closest logarithmic value|
|[*iElements*]|0|Resistor Network Elements<br/>0 = Simple 2 resistor network<br/>1 = Three resistor network, series resistors on bottom<br/>2 = Three resistor network, series resistors on top<br/>3 = Three resistor network, parallel resistors on bottom<br/>4 = Three resistor network, parallel resistors on top|
|[*dMatchType*]|0|Matching when equal ratio offsets <br/>-1 = Choose values that produce Thevenin resistance closest to the middle of the range<br/>0 = Choose first available perfect match<br/>>0 = Choose values to produce a Thevenin resistance closest to dMatchType|
|[*dThevMinLimit*]|1,000|Limit Thevenin resistance to no less than dThevMinLimit|
|[*dThevMaxLimit*]|100,000|Limit Thevenin resistance to no more than dThevMaxLimit|
|[*dCompMinLimit*]|1|Limit any resistor value to no less than dMinLimit|
|[*dCompMaxLimit*]|1,000,000|Limit any resistor value to no more than dMaxLimit|
- [  ]  - Parameters are *optional*



### **Let’s Get Started!**

Before installing the add-in, determine which version of Excel (x32 or x64 bit) you are running by selecting **File**, then **Account** from the main Excel window.   Clicking the **About Excel** button will open a window containing version information at the top.

The add-in is written in Visual Basic .NET, using the free community edition of Visual Studio 2019.  Full source code and compiled .XLL files are available on GitHub: <http://github.com/DonaldSchelle/EETools>.   Download the .XLL for your version of Excel from the **Releases** section on the right-hand side of the page.    The two versions are:

- (x32 bit), EETools-AddIn-packed.xll 
- (x64 bit), EETools-AddIn64-packed.xll 

Follow these steps (Excel 2016) to install the add-in:

1. Click the **File** menu option and select **Options** from the left hand menu.
1. Select the **Add-Ins** tab.  **Manage** the **Excel Add-Ins** (drop-down menu item) by clicking the **Go** button.
1. Click the **Browse** button.  Point to the location of the saved .XLL file and select **OK**.
1. When complete, the add-in is enabled for all spreadsheets.






### **Endless Possibilities**

Once the computational heavy-lifting is automated, many “what-if” scenarios that were previously cumbersome or near-impossible to examine, become trivial.    For example, suppose we want to standardize on a minimal set of resistor values for all designs moving forward.    Minimizing the number of values would reduce the number of component reels needed for the pick-and-place machines during assembly, while leveraging economies of scale due to purchasing higher volumes of fewer component values.

Resistor feedback networks in power regulators are an ideal case for this type of analysis.  Since the feedback resistor values set the nominal output voltage, increasing accuracy of the nominal output voltage maximizes error budget typically reserved for reference and resistor tolerances.   

Suppose that we want to design regulator circuits capable of any output voltage between our power supply feedback reference (typically 0.8V), and 5V, in 1mV increments.   We also want to limit the network Thevenin resistance to between 10kΩ - 100kΩ (fairly standard for power supplies), and use component values in an E6 series limited to between 100Ω - 100kΩ (a scant 19 component values).   Choosing the network during circuit design with the iElements parameter enables a number of networks to optimize the nominal output voltage error yielding some design flexibility.  Nevertheless, this is no doubt a stringent set of design criteria.   Given any desired output voltage, what is the worst-case nominal voltage error (figure 8)?  

<img src="Figures\Figure 8.png" style="zoom:50%;" />

<p align="center"><b>Figure 8</b> – Calculating the best possible error for each output voltage given a constrained set of design criteria becomes trivial once the computational heavy-lifting is automated.</p>



While the worst-case error is 0.5615%, most of the output voltages can be generated with errors less than ~0.3%.   But what if our design criteria changes? Perhaps the achievable error isn’t low enough.   Once the spreadsheet is setup, examining alternate design-criteria takes seconds.    Table 5 details maximum error when given alternate E series (i.e. using more or less component values), and explores an alternate reference voltage of 0.6V, which is increasingly more common in modern voltage regulator ICs.



**Table 5. Maximum Error VOUT from VFB (0.6V/0.8V) – 5.0V in 1mV Increments**


|**E-Series**|**Reference Voltage = 0.6V**|**Reference Voltage = 0.8V**|
| :-: | :-: | :-: |
|E3|6.768749%|6.092789%|
|E6|0.659358%|0.561537%|
|E12|0.141352%|0.134597%|
|E24|0.015328%|0.011362%|
|E48|0.001204%|0.001413%|
|E96|0.000232%|0.000253%|
|E192|0.000000%|0.000000%|



<img src="Figures\Table 5.png" style="zoom: 50%;" />







### **What’s Next?**

There are very few limits regarding automation inside Excel.   Full source-code for all of these functions is published on GitHub and freely available under an MIT license.  The code can be downloaded and used as is, or it can serve as a modifiable template to create your own custom functions.





### **References**

1. Art Kay, Tim Green. “*Analog Engineer's Pocket Reference”, Texas Instruments.* 
   - <http://www.ti.com/seclit/eb/slyw038c/slyw038c.pdf>

2. Christine Schneider. <i>“Excel Formula Calculates Standard 1%-Resistor Values.”</i>, EDN, (2002, January 20)
   - [http://www.electronicdesign.com/technologies/components/article/21763411/excel-formula-calculates-standard-1resistor-values](http://www.electronicdesign.com/technologies/components/article/21763411/excel-formula-calculates-standard-1resistor-values%20) 


3. Donald Schelle.  “*Calculate standard resistor values in Excel”, EDN,* (2013, January 2).
   - <http://www.edn.com/calculate-standard-resistor-values-in-excel/>


4. Dwight Larson. “*Say “No More” to Tedious Calculations for E Series Values*.” (2020, April 7). 
   - <http://www.maximintegrated.com/en/design/blog/say-no-more-to-tedious-calculations-for-e-series-values.html>


5. *Excel-DNA, Free and easy .NET for Excel*. (2020). 
   - <http://excel-dna.net/>

