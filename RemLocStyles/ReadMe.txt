This folder contains layers for "Find XY of Remote Location", ArcGIS tool ("RemoteLocationXY.py" script).
The layers help to set the symbology of the output shp-files:

objectloc.lyr     ----sets--->   ObjectLocation.shp
startpoints.lyr  ----sets--->   ObservationPoints.shp
lines.lyr             ----sets--->   ObservationLines.shp
buffer.lyr          ----sets--->    Accuracy.shp

You need to keep this folder inside the folder of the script for proper work!

If you don't like the curent symbology, you can change it as you wish and then substitute *.lyr-file instead the existing here. So in this case your new lyr-file with the old name will be used as a default pattern for setting of the symbology.