## PeakHeightWidthRatioCalculator_Macro

#### What is this macro for?
This macro imports a file of raw data and carries out a calculation of:  
1. The position of the peak in a chromatographic trace  
2. The width of this peak at half its height

The purpose of this calculation is to check and confirm the results reported by proprietary software provided by the manufacturer of the chromatographic machine.

#### What is in this folder?

*P10.arw, P50.arw, P100.arw and P200.arw* are files which contain the raw data output for high pressure liquid chromatography runs of pullulan molecular weight standards. These standards range from 10k to 200k in molecular weight.

*PeakSeeker.bas* is a text copy of the Visual Basic macro code that is incorporated into the Excel file.

#### To use PeakHeightWidthRatioCalculator_Macro.xls
1. Open *PeakHeightWidthRationCalculator_Macro.xls*
2. Click the pop up that asks if you want to 'Enable Content'.
3. Click the button 'Import Data' and choose one of the .arw data files.
4. Enter peak start and peak end values in cells B4 and B5.
5. Click 'Analyze Peak'.



