# LiftCompleteLite

Lift Complete Lite is an application for managing Olympic Weightlifting
compeititons.

This application is a set of loosely linked apps developed on a base
of Excel's VBA.  It runs on Excel 2007 and later.

Excel's basic functions are completely replaced with a custom "Ribbon"
control set.

Lift Complete Lite manages the flow of compeition from weigh in to 
final score sheets and medals ranking.

This repository contains three applications and documentation.

Lift Weighin  manages the weigh in procedure and creates competition
cards and protocols.  It creates a "Session" file which is stored for
later access.

Lift Manager controls the compeititon and records the results.  Lift
Manager begins by reading the Session file created by Lift Weigh In
saving a lot of typing and eliminating errors.  When in Run Mode, 
Manager then follows the rules of weightlifting to call the next lifter
in the competition to the platform.

Lift Loader is a graphical depiction of the required weights on the bar.
This program shows the loaders which weights are needed for the bar to 
be properly loaded on the platform.  Loader reads a file that is saved
on the Manager system so that it automatically updates.

For best results, all of these systems should be on the same LAN.

Lift Complete Lite runs on Excel 2007 and higher on Windows 7 and later.