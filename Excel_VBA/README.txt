PeakHeightWidthRatioCalculator_Macro

What is this macro for?

This macro imports a file of raw data and does a calculation of:
a. the position of the peak in a chromatographic trace 
b. the width of this peak at half its height

The purpose of this calculation is to check and confirm the results reported by proprietary software provided by the manufacturer of the chromatographic machine.

What is in this folder?

P10.arw, P50.arw, P100.arw and P200.arw are files which contain the raw data output for high pressure liquid chromatography runs of Pullulan molecular weight standards. These standards range from 10k to 200k in Molecualr weight.

PeakSeeker.bas is a text copy of the Visual Basic macro code that is incorporated into the Excel file.

To use PeakHeightWidthRatioCalculator_Macro.xls, 

1. Open the Excel file, click the pop up that asks if you want to 'Enable Content'.
2. Click the button 'Import Data' and choose one of the .arw data files.
3. Examine the chart and enter peak start and peak end values in cells B4 and B5
4. Click 'Analyze Peak'


