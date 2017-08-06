# XRDCalculator
Python script for calculating Scherrer Widths, d-spacings, and plotting diffraction patterns from XRD spreadsheet.

Calculations (require user input)
  - calculates Scherrer widths given user input of 2Theta peak position and instrument parameters (x-ray radiation wavelength, line broadening factors)
  - calculates average Scherrer width of pattern given numerous Scherrer widths
  - calculates d-spacings of peaks given user input of 2Theta peak positions
  
Plotting function utilizes Excel spreadsheet with 2Theta values in B-column and intensity values in C-column for plot
  - plots diffraction pattern with user input for chart titles, axes labels, and 2Theta range. Supports multiple diffraction patterns on a single plot
  
