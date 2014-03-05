'''**************************************************************************
"Triangulate the XY of Remote Location" is a simple tool
for ArcGIS, which finds the coordinates of a remote object using triangulation data.
The triangulation data should include azimuths to an observed object from three
points of observation and the coordinates of the observation points (in WGS84
decimal degrees). The tool estimates an error of the measurements and
displays it as a buffer of location accuracy.

The input and output data are stored in Excel spreadsheet (.xls) and a Comma Separated
Value text file (.txt). The name and folder of the output files could be choosen by user
or left in the default folder (C:\Temp_RemLocXY\RemoteLocationXY). 

The script bellow (RemoteLocationXY.py) was developed by Dorn Moore, Spatial
Analyst and Dmitrii Sarichev, GIS Intern at the International Crane Foundation
in January-February 2014

For more information cotact dorn@savingcranes.org.
**************************************************************************'''

arcpy.AddWarning("Triangulate the XY of Remote Location\nCopyright (C) 2014  International Crane Foundation\n")
arcpy.AddWarning("This program is free software: you can redistribute it and or modify\nit under the terms of the GNU General Public License as published by \nthe Free Software Foundation, either version 3 of the License, or\nany later version.\n")
arcpy.AddWarning("This program is distributed in the hope that it will be useful,\nbut WITHOUT ANY WARRANTY; without even the implied warranty of \nMERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the \nGNU General Public License for more details. \n\nYou should have received a copy of the GNU General Public License\nalong with this program.  If not, see <http://www.gnu.org/licenses/>.\n")

# Import modules
import arcpy, os, sys, csv, math
from arcpy import env

#----------------------------------------------------------------------------
#                            Workspace
#----------------------------------------------------------------------------

# Choose path and folder or create "C:\Temp_RemLocXY" as a workspace by default
TempDir = arcpy.GetParameterAsText(11)
if TempDir == '#' or not TempDir: 
    if not os.path.exists("C:\\Temp_RemLocXY"):  # provide the
        os.makedirs("C:\\Temp_RemLocXY")         # default path
    TempDir = "C:\\Temp_RemLocXY"                # if unspecified
    
# Set current workspace
arcpy.env.scratchWorkspace = TempDir         
arcpy.env.workspace = TempDir
# Define Spatial Reference for WGS84
wgs = arcpy.SpatialReference(4326)

#----------------------------------------------------------------------------
#                   Create Txt-File and Input Data
#----------------------------------------------------------------------------

# Input name of the file or set "RemoteLocationXY" by default
name = arcpy.GetParameterAsText(12)
if name == '#' or not name: 
    name = "RemoteLocationXY"
filename = name + ".txt"
filepath = os.path.join(TempDir, filename)   # Provide filepath

# Make the first row a list of the field names
fields = ["lat1","lon1","az1","lat2","lon2","az2","lat3","lon3","az3","dist","Xin","Yin","r"]

# Get the parameters via user input
lat1 = arcpy.GetParameterAsText(0)
lon1 = arcpy.GetParameterAsText(1)
az1 = arcpy.GetParameterAsText(2)
lat2 = arcpy.GetParameterAsText(3)
lon2 = arcpy.GetParameterAsText(4)
az2 = arcpy.GetParameterAsText(5)
lat3 = arcpy.GetParameterAsText(6)
lon3 = arcpy.GetParameterAsText(7)
az3 = arcpy.GetParameterAsText(8)
dist = arcpy.GetParameterAsText(9)
table = arcpy.SetParameter(10, filepath)

# Create average (avg) latitude and longitude for use later
avgLat = (float(lat1)+float(lat2)+float(lat3))/3
avgLon = (float(lon1)+float(lon2)+float(lon3))/3

# Set the values as nought by default
Xin = 0
Yin = 0
r = 0

# Make the second row a list of the parameters
row_input = [lat1,lon1,az1,lat2,lon2,az2,lat3,lon3,az3,dist,Xin,Yin,r]
 
# Write the CSV file
with open(filepath,"wb") as f:
        writer = csv.writer(f)     
        write = writer.writerow
        write(fields)
        write(row_input)
f.close()

#----------------------------------------------------------------------------
#                  Overwrite Existing Maplayers
#----------------------------------------------------------------------------

# Enable the ability to overwrite existing data
arcpy.env.overwriteOutput = True

# Get the map document
mxd = arcpy.mapping.MapDocument("CURRENT")

# Get the data frame
df = arcpy.mapping.ListDataFrames(mxd,"*")[0]

# Clean up old layers created during previous Visualizations
##For each layer in dataframe of the current mxd file
for lyr in arcpy.mapping.ListLayers(mxd, "", df):
    try:
        # If the workspacePath is equal to our folder path
        if lyr.workspacePath == TempDir:
            #Delete the layer so we start fresh
            arcpy.mapping.RemoveLayer(df, lyr)
    except:
        arcpy.AddWarning("Web based layer left on map.")
     
#----------------------------------------------------------------------------
#         Create and Display Lines and Points of Observations
#----------------------------------------------------------------------------

# Set local variables
## Were using a Concatenate to make the path of the folderpath
## This combines the filename and filepath so the OS can read it.
line_1 = os.path.join(str(TempDir),"line_1.shp")
line_2 = os.path.join(str(TempDir),"line_2.shp")
line_3 = os.path.join(str(TempDir),"line_3.shp")
lines = os.path.join(str(TempDir),"ObservationLines.shp")
startpoint_1 = os.path.join(str(TempDir),"ObsPoint1")
startpoint_2 = os.path.join(str(TempDir),"ObsPoint2")
startpoint_3 = os.path.join(str(TempDir),"ObsPoint3")
startpoints = os.path.join(str(TempDir),"ObservationPoints.shp")

# Create lines and points
## Process: bearings and distances to lines
arcpy.management.BearingDistanceToLine(filepath, line_1, "lon1", "lat1", "dist", "METERS", "az1", "DEGREES", "GEODESIC", "", wgs)
arcpy.management.BearingDistanceToLine(filepath, line_2, "lon2", "lat2", "dist", "METERS", "az2", "DEGREES", "GEODESIC", "", wgs)
arcpy.management.BearingDistanceToLine(filepath, line_3, "lon3", "lat3", "dist", "METERS", "az3", "DEGREES", "GEODESIC", "", wgs)

## Make XY Event Layers: creating of start points in WGS84 using inputed data
arcpy.management.MakeXYEventLayer(filepath, "lon1", "lat1", startpoint_1, wgs, "")
arcpy.management.MakeXYEventLayer(filepath, "lon2", "lat2", startpoint_2, wgs, "")
arcpy.management.MakeXYEventLayer(filepath, "lon3", "lat3", startpoint_3, wgs, "")

## Merge all points and all lines in two layers by geometry type
arcpy.management.Merge([startpoint_1,startpoint_2,startpoint_3],startpoints)
arcpy.management.Merge([line_1,line_2,line_3],lines)

# Change tables of the new layers
## Delete fields excluding "dist" field and use it as the field of "distingush" for setting of unique symbology
## Calculate ID of "dist" fields
arcpy.management.DeleteField(lines,["lat1","lon1","az1","lat2","lon2","az2","lat3","lon3","az3"])
arcpy.management.CalculateField(lines, "dist", "!FID!+1", "PYTHON", "")
arcpy.management.DeleteField(startpoints,["lat1","lon1","az1","lat2","lon2","az2","lat3","lon3","az3","Xin","Yin","r"])
arcpy.management.CalculateField(startpoints, "dist", "!FID!+1", "PYTHON", "")

# Adding outputs to the map
## create new layers
newlayer0 = arcpy.mapping.Layer(lines)
newlayer1 = arcpy.mapping.Layer(startpoints)

## add the layer to the map at the bottom of the TOP in data frame 0
arcpy.mapping.AddLayer(df, newlayer0,"TOP")
arcpy.mapping.AddLayer(df, newlayer1,"TOP")

## Get relative path of the executing script. It is required for creating path to the templates (*.lyr-files) of symbology   
relatpath = os.path.dirname(__file__) # in case of error here you can use alternative way: os.path.realpath(os.path.dirname(sys.argv[0]))

## Set symbology of the buffer according to the "startpoints.lyr" template
try:
    ### Change simbology of the points
    updateLayer = arcpy.mapping.ListLayers(mxd, "ObservationPoints", df)[0]
    stylepath = relatpath+"\\RemLocStyles\\startpoints.lyr"  # if you do not like the default symbology of "ObservationPoints" layer you can change this lyr-file as you wish and it will be used as a new default template
    sourceLayer = arcpy.mapping.Layer(stylepath)
    arcpy.mapping.UpdateLayer(df, updateLayer, sourceLayer, True)
    updateLayer.showLabels = True # switching on of the labels
    ##### Set label properties
    if updateLayer.supports("LABELCLASSES"):
        for lblClass in updateLayer.labelClasses:
            lblClass.expression = '"<CLR red=\'178\' green=\'178\' blue=\'178\'><FNT size = \'9\'>"&[dist]&"</FNT></CLR>"' # here you can change color and size of the labels
            lblClass.showLabels = True
except:
    ### if the previous block of code hasn't changed the symbology it means the patterns of the layers haven't been found
    arcpy.AddWarning("\n*.lyr-pattern of the \"ObservationPoints\" layer hasn't been found.\n  The symbology has been set by default.\n")

## Set symbology of the startpoints layer according to the "lines.lyr" template
try:
    ### Change simbology of the lines
    updateLayer = arcpy.mapping.ListLayers(mxd, "ObservationLines", df)[0]
    stylepath = relatpath+"\\RemLocStyles\\lines.lyr" # if you do not like the default symbology of "ObservationLines" layer you can change this lyr-file as you wish and it will be used as a new default template
    sourceLayer = arcpy.mapping.Layer(stylepath)
    arcpy.mapping.UpdateLayer(df, updateLayer, sourceLayer, True)
except:
    ### if the previous block of code hasn't changed the symbology it means the patterns of the layers haven't been found
    arcpy.AddWarning("\n*.lyr-pattern of the \"ObservationLines\" layer hasn't been found.\n  The symbology has been set by default.\n")

#----------------------------------------------------------------------------
#          Get Points of Intersection and Check Triangulation
#----------------------------------------------------------------------------

# Find intersections of line_1, line_2 and line_3
## Local variables
XY1 = os.path.join(str(TempDir),"RemLocXYTemp_XY1.shp")
XY2 = os.path.join(str(TempDir),"RemLocXYTemp_XY2.shp")
XY3 = os.path.join(str(TempDir),"RemLocXYTemp_XY3.shp")

## Try to create the intersection point of lines 2 and 3
arcpy.analysis.Intersect(line_2 + " #;" + line_3 +" #", XY1, "ALL", "", "POINT")
### Count the points in the output layer and export the result in integer value
Count1 = arcpy.management.GetCount(XY1)
CountXY1 = int(Count1.getOutput(0))
### Make sure, there is only one point in the output layer
if CountXY1!=1:   
    # If it is false then print the warning:
    arcpy.AddWarning("\nObservation lines #2 and #3 do not intersect.\nThe triangulation is invalid!\n")
    
## Try to create the intersection point of lines 1 and 3
arcpy.analysis.Intersect(line_1 + " #;" + line_3 +" #", XY2, "ALL", "", "POINT")
### Count the points in the output layer and export the result in integer value
Count2 = arcpy.management.GetCount(XY2)
CountXY2 = int(Count2.getOutput(0))
### Make sure, there is only one point in the output layer
if CountXY2!=1:
    # If it is false then print the warning:
    arcpy.AddWarning("\nObservation lines #1 and #3 do not intersect.\nThe triangulation is invalid!\n")
    
## Try to create the intersection point of lines 1 and 2
arcpy.analysis.Intersect(line_1 + " #;" + line_2 +" #", XY3, "ALL", "", "POINT")
### Count the points in the output layer and export the result in integer value
Count3 = arcpy.management.GetCount(XY3)
CountXY3 = int(Count3.getOutput(0))
### Make sure, there is only one point in the output layer
if CountXY3!=1:
    # If it is false then print the warning:
    arcpy.AddWarning("\nObservation lines #1 and #2 do not intersect.\nThe triangulation is invalid!\n")

#----------------------------------------------------------------------------
# If There is a Triangle then find The Center and The Radius of The Incircle
#----------------------------------------------------------------------------

# If there are all three points then execute triangulation: find incenter and error of the measurements 
IntersectionCount = CountXY1+CountXY2+CountXY3
if IntersectionCount == 3:
    # Calculate incenter coordinates
    ## Reprogect point layers in UTM
    ### Local variable
    XY1_UTM = os.path.join(str(TempDir),"RemLocXYTemp_XY1_UTM.shp")
    XY2_UTM = os.path.join(str(TempDir),"RemLocXYTemp_XY2_UTM.shp")
    XY3_UTM = os.path.join(str(TempDir),"RemLocXYTemp_XY3_UTM.shp")

    ### Determine the correct UTM Zone
    if avgLat >0:   #Northern Hemisphere
        prj = arcpy.SpatialReference(int("326" + str(int(math.floor((avgLon + 180)/6) + 1))))
    else:           #Southern Hemisphere
        prj = arcpy.SpatialReference(int("327" + str(int(math.floor((avgLon + 180)/6) + 1))))
    
    ### Reprojection:
    arcpy.management.Project(XY1, XY1_UTM, prj)
    arcpy.management.Project(XY2, XY2_UTM, prj)
    arcpy.management.Project(XY3, XY3_UTM, prj)

    ## Define variable for incenter calculation
    ### Get X1, Y1 coordinates
    desc = arcpy.Describe(XY1_UTM)
    shapefieldname = desc.ShapeFieldName
    row = arcpy.SearchCursor(XY1_UTM)
    for row in row:
        feat = row.getValue(shapefieldname)
        startpt = feat.firstPoint
        X1 = startpt.X
        Y1 = startpt.Y
        
    ### Get X2, Y2 coordinates
    desc = arcpy.Describe(XY2_UTM)
    shapefieldname = desc.ShapeFieldName
    row = arcpy.SearchCursor(XY2_UTM)
    for row in row:
        feat = row.getValue(shapefieldname)
        startpt = feat.firstPoint
        X2 = startpt.X
        Y2 = startpt.Y
        
    ### Get X3, Y3 coordinates
    desc = arcpy.Describe(XY3_UTM)
    shapefieldname = desc.ShapeFieldName
    row = arcpy.SearchCursor(XY3_UTM)
    for row in row:
        feat = row.getValue(shapefieldname)
        startpt = feat.firstPoint
        X3 = startpt.X
        Y3 = startpt.Y
        
    ### Calculate the side A of the triangle (X2Y2 <-----> X3Y3)
    A = math.sqrt((X2-X3)*(X2-X3)+(Y2-Y3)*(Y2-Y3))
    ### Calculate the side B of the triangle (X1Y1 <-----> X3Y3)
    B = math.sqrt((X1-X3)*(X1-X3)+(Y1-Y3)*(Y1-Y3))    
    ### Calculate the side C of the triangle (X1Y1 <-----> X2Y2)
    C = math.sqrt((X1-X2)*(X1-X2)+(Y1-Y2)*(Y1-Y2))
    
    ## Calculate the X Y coordinate of the incenter
    Xin_UTM = ((A*X1)+(B*X2)+(C*X3))/(A+B+C)
    Yin_UTM = ((A*Y1)+(B*Y2)+(C*Y3))/(A+B+C)
      
    # Calculate the error of the measuremet (it equals R of incircle)
    ## Find semiperimeter: p = (A+B+C)/2
    p = (A+B+C)/2
    ## Calculate area of the triangle: S = sqroot(p*(p-A)*(p-B)*(p-C))
    S = math.sqrt(p*(p-A)*(p-B)*(p-C))
    ## Calculate error: R = S/p
    R = S/p
    
    # Input the coordinates of incenter and the value of error in the temporary txt-file
    ## Set variable
    filepath_UTM = os.path.join(TempDir,"RemLocXYTempUTM.txt")
    row_intermed = [lat1,lon1,az1,lat2,lon2,az2,lat3,lon3,az3,dist,Xin_UTM,Yin_UTM,R]
    ## Write data into the temporary txt-file
    with open(filepath_UTM,"wb") as f:
        writer = csv.writer(f)     
        write = writer.writerow
        write(fields)
        write(row_intermed)
    f.close()
  
    # Make XY event layer
    Incenter_UTM = os.path.join(str(TempDir),"ObjectsLocation_UTM.shp")
    arcpy.management.MakeXYEventLayer(filepath_UTM, "xin", "yin", Incenter_UTM, prj)
    
    ## Reproject the incenter layer back to WGS 84
    Incenter = os.path.join(str(TempDir),"ObjectLocation.shp") 
    arcpy.management.Project(Incenter_UTM, Incenter, wgs)

    ## Add the new layer in the current map frame
    newlayer2 = arcpy.mapping.Layer(Incenter)
    arcpy.mapping.AddLayer(df, newlayer2,"TOP")
    ### Set the symbology of the "ObjectLocation" layer according to the "objectloc.lyr" template
    try:
        #### Change symbology of the points
        updateLayer = arcpy.mapping.ListLayers(mxd, "ObjectLocation", df)[0]
        stylepath = relatpath+"\\RemLocStyles\\objectloc.lyr" # if you do not like the default symbology you can change this lyr-file as you wish and it will be used as a new default template
        sourceLayer = arcpy.mapping.Layer(stylepath)
        arcpy.mapping.UpdateLayer(df, updateLayer, sourceLayer, True)
        updateLayer.showLabels = True
        #### Set label properties
        if updateLayer.supports("LABELCLASSES"):
            for lblClass in updateLayer.labelClasses:
                lblClass.expression = '"<CLR red=\'178\' green=\'178\' blue=\'178\'><FNT size = \'9\'>" & "X "& [Xin] &vbCrLf&"Y "& [Yin] &vbCrLf& "Error = "& round( [r], 1)&" m" & "</FNT></CLR>"'
                lblClass.showLabels = True
    except:
        #### if the previous block of code hasn't changed the symbology it means the patterns of the layers haven't been found
        arcpy.AddWarning("\n*.lyr-pattern of the \"ObjectLocation\" layer hasn't been found.\n  The symbology has been set by default.\n")
            
    ## Assign variables:
    ### Get Xin, Yin coordinates
    desc = arcpy.Describe(Incenter)
    shapefieldname = desc.ShapeFieldName
    row = arcpy.SearchCursor(Incenter)
    for row in row:
        feat = row.getValue(shapefieldname)
        startpt = feat.firstPoint
        Xin = startpt.X
        Yin = startpt.Y
        
    ## Rewrite the coordinates of incenter in the inputed txt-file
    row_output = [lat1,lon1,az1,lat2,lon2,az2,lat3,lon3,az3,dist,Xin,Yin,R]
    with open(filepath,"wb") as f:
        writer = csv.writer(f)     
        write = writer.writerow
        write(fields)
        write(row_output)
    f.close()
    
    # Calculate field: returning of WGS84 coordinates of the center into the attribute table of ObjectLocation.shp
    arcpy.management.CalculateField(Incenter, "Xin", Xin, "PYTHON")
    arcpy.management.CalculateField(Incenter, "Yin", Yin, "PYTHON")
    
    # Export txt-file to Excel table
    name_xls = name + '.xls'
    output_xls = os.path.join(TempDir, name_xls)
    arcpy.conversion.TableToExcel(Incenter, output_xls)

    ## Add result message in processing box
    arcpy.AddMessage("\nObject Location: \n" + "X " + str(Xin) + ", Y " + str(Yin) + " (in WGS-84)" + "\nEst. Error = " + str(R) + " meters \n")

    # Create buffer of errors around of the incenter; radius of the buffer equal the radius of the incircle (R)
    bufferror = os.path.join(str(TempDir),"Accuracy.shp")
    arcpy.analysis.Buffer(Incenter, bufferror, str(R) + " Meters", "FULL", "ROUND", "NONE", "")
    ## Add new buffer layer in the current map frame
    newlayer3 = arcpy.mapping.Layer(bufferror)
    arcpy.mapping.AddLayer(df, newlayer3,"AUTO_ARRANGE")
    ### Set symbology of the buffer according to the "buffer.lyr" template
    try:
        #### Change simbology of the buffer
        updateLayer = arcpy.mapping.ListLayers(mxd, "Accuracy", df)[0]
        stylepath = relatpath+"\\RemLocStyles\\buffer.lyr" # if you do not like the default symbology you can change this lyr-file as you wish and it will be used as a new default template
        sourceLayer = arcpy.mapping.Layer(stylepath)
        arcpy.mapping.UpdateLayer(df, updateLayer, sourceLayer, True)
        updateLayer.transparency = 70 # set transparency of the "Accuracy" layer. It is 70% by default
    except:
        #### if the previous block of code hasn't changed the symbology it means the patterns of the layers haven't been found
        arcpy.AddWarning("\n*.lyr-pattern of the \"Accuracy\" layer hasn't been found.\n  The symbology has been set by default.\n")
    
    # Delete intermediate files
    ## Check to see if intermediate data exist; if they do, then delete them
    intermed = [line_1,line_2,line_3,startpoint_1,startpoint_2,startpoint_3,XY1,XY2,XY3,XY1_UTM,XY2_UTM,XY3_UTM,filepath_UTM,Incenter_UTM]
    for intermed in intermed:
        if arcpy.Exists(intermed):
            arcpy.management.Delete(intermed)

else:
    # Print error message in processing box:
    arcpy.AddWarning("There is insufficient or wrong data for build of triangle, deriving coordinates\nof observed object and estimation of error of the measurements\n")
    
    # Delete intermediate files
    ## Check to see if intermediate data exist; if they do, then delete them
    intermed = [line_1,line_2,line_3,startpoint_1,startpoint_2,startpoint_3,XY1,XY2,XY3]
    for intermed in intermed:
        if arcpy.Exists(intermed):
            arcpy.management.Delete(intermed)
