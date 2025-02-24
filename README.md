# Simulation and Curve Fitting in Excel

Work in progress, not ready to use!

## Motivation

Actuaries do curve fitting and Monte Carlo simulations in R or Python for developing insurance policy pricing.
However, we often need to make these tools available to others who do not use R or Python, e.g., underwriters.
Connecting R and Python is possible but always feels brittle and hard to maintain.
Why can't Excel do these things closer to the spreadsheet?
Spreadsheet functions are out of the question, but what about Javascript via Office 365 add-ins?  Or VBA?!

## Goal

Can I build a system in VBA that 

1. Replicates common curve fitting and simulation pricing done in R
2. Giving the exact same random numbers when provided the same seed (for reproducibility and testing)
3. Is fast enough

