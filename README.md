# google-spreadsheets-whatif
Script to emulate the What-If feature of Microsoft Excel

Usage:

=WhatIf("C3", "B1", A1:A3)

This example takes the formula in C3 cell and replaces the value of the parameter in B1 within that formula with the values in the range A1:A3.
Results are shown vertically below the cell containing the WhatIf call, including it.