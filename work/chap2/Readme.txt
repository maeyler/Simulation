
Excel Examples for Chapter 2, Discrete-Event System Simulation, 4th edition:

Version 1, Dec 21, 2004

The examples have been tested in Excel 2000 (SP-3) and Excel 2003 (SP1), and get the same results in both versions.

In these versions of Excel, you should get the same results as the textbook, provided you do the following:

1.  In either worksheet, set the Seed to a value of 12345.
2.  Click on the "Reset Seed" button, which does 2 things.  It sets the seed to the given value and runs the simulation once.

Note:  The "Generate New Trial" and "Run Experiment" buttons do NOT set the seed.  Each time they are clicked, they use new random numbers by simply continuing down the random "stream" in the built-in random number generator.

The Excel examples do not use Excel's random number generator in macro formulas; instead, they call VB functions which use VB's random number generator.   If a later version of Excel uses a different random number generator for VBA (Visual Basic), then most likely the results will change and no longer match the textbook.  

