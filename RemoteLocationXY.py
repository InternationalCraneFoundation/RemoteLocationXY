'''**************************************************************************
This is a simple tool for ArcGIS, which finds coordinates of a remote object
using triangulation data. The triangulation data should include azimuths
to an observed object from three points of observation and the coordinates
of the observation points (in WGS84 decimal degrees). The tool allows to
estimate an error of the measurements and displays it as a buffer of location
accuracy.

The input and output data would be stored in xls-spreadsheet and txt-file.
The name and folder of the output files could be choosen by user or stayed
by default (C:\Temp_TriVis\TriVisTemp) 

The script bellow was developed by Dorn Moore, Spatial Analyst and
Dmitrii Sarichev, GIS Intern at the International Crane Foundation
in January-February 2014
**************************************************************************'''

# Import modules
import arcpy, os, csv, math
from arcpy import env

#------------------------------------------------------------------------
#                            Workspace
#------------------------------------------------------------------------

# Choose path and folder or create "C:\\temp" as a workspace by default
TriVisDir = arcpy.GetParameterAsText(11)
if TriVisDir == '#' or not TriVisDir: 
    if not os.path.exists("C:\\Temp_TriVis"):  # provide the
        os.makedirs("C:\\Temp_TriVis")         # difault path
    TriVisDir = "C:\\Temp_TriVis"              # if unspecified
    
# Set current workspace
arcpy.env.scratchWorkspace = TriVisDir         
arcpy.env.workspace = TriVisDir

#------------------------------------------------------------------------
#                   Create Txt-File and Input Data
#------------------------------------------------------------------------

# Input name of the file or set "TriVisTemp.txt" by default
name = arcpy.GetParameterAsText(12)
if name == '#' or not name: 
    name = "TriVisTemp"
filename = name + ".txt"
filepath = os.path.join(TriVisDir, filename)   # Provide filepathe

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

# Create AVG lat and lon for use later
avgLat = (float(lat1)+float(lat2)+float(lat3))/3
avgLon = (float(lon1)+float(lon2)+float(lon3))/3

# Define Spatial Refernence for WGS84
wgs = arcpy.SpatialReference(4326)

# Set the values as nought by default
Xin = 0
Yin = 0
r = 0

# Make the second row a list of the parameters
row_input = [lat1,lon1,az1,lat2,lon2,az2,lat3,lon3,az3,dist,Xin,Yin,r]
 
# Write the CSV file
writer = csv.writer(open(filepath,"wb"))
write = writer.writerow
write(fields)
write(row_input)

with open(filepath,"wb") as f:
        writer = csv.writer(f)     
        write = writer.writerow
        write(fields)
        write(row_input)
f.close()

#------------------------------------------------------------------------
#                  Overwriting of Inputted Data
#------------------------------------------------------------------------

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
        if lyr.workspacePath == TriVisDir:
            #Delete the layer so we start fresh
            arcpy.mapping.RemoveLayer(df, lyr)
    except:
        arcpy.AddWarning("Removal of web based layer skipped.")
     
#------------------------------------------------------------------------
#         Create and Display Lines and Points of Observations
#------------------------------------------------------------------------

# Local variables:
## Were using a Concatenate to make the path of the folderpath
## This combines the filename and filepath so the OS can read it.
Line_3 = os.path.join(str(TriVisDir),"TriVisTemp_BearingDistanceTo3.shp")
Line_1 = os.path.join(str(TriVisDir),"TriVisTemp_BearingDistanceTo1.shp")
Line_2 = os.path.join(str(TriVisDir),"TriVisTemp_BearingDistanceTo2.shp")
start_point_1 = "TriVisTemp_Layer1"
start_point_2 = "TriVisTemp_Layer2"
start_point_3 = "TriVisTemp_Layer3"
points_shp = os.path.join(str(TriVisDir),"TriangulationPoints.shp")

# Create lines and points
## Process: bearings and distances to lines
arcpy.BearingDistanceToLine_management(filepath, Line_1, "lon1", "lat1", "dist", "METERS", "az1", "DEGREES", "GEODESIC", "", "GEOGCS['GCS_WGS_1984',DATUM['D_WGS_1984',SPHEROID['WGS_1984',6378137.0,298.257223563]],PRIMEM['Greenwich',0.0],UNIT['Degree',0.0174532925199433]];-400 -400 1000000000;-100000 10000;-100000 10000;8.98315284119522E-09;0.001;0.001;IsHighPrecision")
arcpy.BearingDistanceToLine_management(filepath, Line_2, "lon2", "lat2", "dist", "METERS", "az2", "DEGREES", "GEODESIC", "", "GEOGCS['GCS_WGS_1984',DATUM['D_WGS_1984',SPHEROID['WGS_1984',6378137.0,298.257223563]],PRIMEM['Greenwich',0.0],UNIT['Degree',0.0174532925199433]];-400 -400 1000000000;-100000 10000;-100000 10000;8.98315284119522E-09;0.001;0.001;IsHighPrecision")
arcpy.BearingDistanceToLine_management(filepath, Line_3, "lon3", "lat3", "dist", "METERS", "az3", "DEGREES", "GEODESIC", "", "GEOGCS['GCS_WGS_1984',DATUM['D_WGS_1984',SPHEROID['WGS_1984',6378137.0,298.257223563]],PRIMEM['Greenwich',0.0],UNIT['Degree',0.0174532925199433]];-400 -400 1000000000;-100000 10000;-100000 10000;8.98315284119522E-09;0.001;0.001;IsHighPrecision")

## Make XY Event Layers
arcpy.MakeXYEventLayer_management(filepath, "lon1", "lat1", start_point_1, "", "")
arcpy.MakeXYEventLayer_management(filepath, "lon2", "lat2", start_point_2, "", "")
arcpy.MakeXYEventLayer_management(filepath, "lon3", "lat3", start_point_3, "", "")

## Process: AddXY Event Layers to single Point file
arcpy.CopyFeatures_management(start_point_1, points_shp)
arcpy.Append_management(start_point_2, points_shp)
arcpy.Append_management(start_point_3, points_shp)

# Adding outputs to map
# create a new layer
newlayer0 = arcpy.mapping.Layer(Line_1)
newlayer1 = arcpy.mapping.Layer(Line_2)
newlayer2 = arcpy.mapping.Layer(Line_3)
newlayer3 = arcpy.mapping.Layer(points_shp)

# add the layer to the map at the bottom of the TOC in data frame 0
arcpy.mapping.AddLayer(df, newlayer0,"TOP")
arcpy.mapping.AddLayer(df, newlayer1,"TOP")
arcpy.mapping.AddLayer(df, newlayer2,"TOP")
arcpy.mapping.AddLayer(df, newlayer3,"TOP")

#------------------------------------------------------------------------
#          Get Points of Intersection and Check Triangulation
#------------------------------------------------------------------------

# Find intersections of Line_1, Line_2 and Line_3
## Local variables
XY1 = os.path.join(str(TriVisDir),"TriVisTemp_XY1.shp")
XY2 = os.path.join(str(TriVisDir),"TriVisTemp_XY2.shp")
XY3 = os.path.join(str(TriVisDir),"TriVisTemp_XY3.shp")

## Try to create the intersection point of lines 2 and 3
arcpy.Intersect_analysis(Line_2 + " #;" + Line_3 +" #", XY1, "ALL", "", "POINT")
### Count the points in the output layer and export the result in integer value
Count1 = arcpy.GetCount_management(XY1)
CountXY1 = int(Count1.getOutput(0))
### Make sure, there is only one point in the output layer
if CountXY1!=1:   
    # If it is false then print the warning:
    arcpy.AddWarning("\nObservation lines #2 and #3 do not intersect.\nThe triangulation is invalid!\n")
    
## Try to create the intersection point of lines 1 and 3
arcpy.Intersect_analysis(Line_1 + " #;" + Line_3 +" #", XY2, "ALL", "", "POINT")
### Count the points in the output layer and export the result in integer value
Count2 = arcpy.GetCount_management(XY2)
CountXY2 = int(Count2.getOutput(0))
### Make sure, there is only one point in the output layer
if CountXY2!=1:
    # If it is false then print the warning:
    arcpy.AddWarning("\nObservation lines #1 and #3 do not intersect.\nThe triangulation is invalid!\n")
    
## Try to create the intersection point of lines 1 and 2
arcpy.Intersect_analysis(Line_1 + " #;" + Line_2 +" #", XY3, "ALL", "", "POINT")
### Count the points in the output layer and export the result in integer value
Count3 = arcpy.GetCount_management(XY3)
CountXY3 = int(Count3.getOutput(0))
### Make sure, there is only one point in the output layer
if CountXY3!=1:
    # If it is false then print the warning:
    arcpy.AddWarning("\nObservation lines #1 and #2 do not intersect.\nThe triangulation is invalid!\n")

#------------------------------------------------------------------------
#     If There is a Triangle then find Center and Radius of Incircle
#------------------------------------------------------------------------

# If there are all three points then execute triangulation: find incenter and error of measurements 
IntersectionCount = CountXY1+CountXY2+CountXY3
if IntersectionCount == 3:
    # Calculate incenter coordinates
    ## Reprogect point layers in UTM (zone 16N)
    ### Local variable
    XY1_UTM = os.path.join(str(TriVisDir),"TriVisTemp_XY1_UTM.shp")
    XY2_UTM = os.path.join(str(TriVisDir),"TriVisTemp_XY2_UTM.shp")
    XY3_UTM = os.path.join(str(TriVisDir),"TriVisTemp_XY3_UTM.shp")

    ### Determine the correct UTM Zone
    if avgLat >0:   #Northern Hemisphere
        prj = arcpy.SpatialReference(int("326" + str(int(math.floor((avgLon + 180)/6) + 1))))
    else:           #Southern Hemisphere
        prj = arcpy.SpatialReference(int("327" + str(int(math.floor((avgLon + 180)/6) + 1))))
    
    ### Reprojection:
    arcpy.Project_management(XY1, XY1_UTM, prj)
    arcpy.Project_management(XY2, XY2_UTM, prj)
    arcpy.Project_management(XY3, XY3_UTM, prj)

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
    ## Calculate error: R = S/p ("//" - return integer)
    R = S/p
    
    # Input the coordinates of incenter and the value of error in the inputed txt-file
    filepath_UTM = os.path.join(TriVisDir,"TriVisTempUTM.txt")
    row_intermed = [lat1,lon1,az1,lat2,lon2,az2,lat3,lon3,az3,dist,Xin_UTM,Yin_UTM,R]

    with open(filepath_UTM,"wb") as f:
        writer = csv.writer(f)     
        write = writer.writerow
        write(fields)
        write(row_intermed)
    f.close()
  
    # Make XY event layer
    Incenter_UTM = os.path.join(str(TriVisDir),"ObjectsLocation_UTM.shp")
    arcpy.MakeXYEventLayer_management(filepath_UTM, "xin", "yin", Incenter_UTM, prj)
    
    ## Reproject the incenter layer back to WGS 84
    Incenter = os.path.join(str(TriVisDir),"ObjectLocation.shp") 
    arcpy.Project_management(Incenter_UTM, Incenter, wgs)
    newlayer5 = arcpy.mapping.Layer(Incenter)
    arcpy.mapping.AddLayer(df, newlayer5,"TOP")
    
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
    arcpy.CalculateField_management(Incenter, "Xin", Xin, "PYTHON")
    arcpy.CalculateField_management(Incenter, "Yin", Yin, "PYTHON")
    
    # Export txt-file to Excel table
    name_xls = name + '.xls'
    output_xls = os.path.join(TriVisDir, name_xls)
    arcpy.TableToExcel_conversion(Incenter, output_xls)

    ## Add result message in processing box
    arcpy.AddMessage("\nObject Location: \n" + "X " + str(Xin) + ", Y " + str(Yin) + " (in WGS-84)" + "\nEst. Error = " + str(R) + " m \n")

    # Create buffer of errors around of incenter; radius of the buffer equal the radius of the incircle (R)
    bufferror = os.path.join(str(TriVisDir),"Accuracy.shp")
    arcpy.Buffer_analysis(Incenter, bufferror, str(R) + " Meters", "FULL", "ROUND", "NONE", "")
    newlayer6 = arcpy.mapping.Layer(bufferror)
    arcpy.mapping.AddLayer(df, newlayer6,"AUTO_ARRANGE")
    try:
        ## Change simbology of the buffer
        updateLayer = arcpy.mapping.ListLayers(mxd, "Accuracy", df)[0]
        stylepath = r"C:\Users\Dmitrii\Desktop\NewTriVisScript\Style\Buffer.lyr" # !!! Need use relative path!!!
        sourceLayer = arcpy.mapping.Layer(stylepath)
        arcpy.mapping.UpdateLayer(df, updateLayer, sourceLayer, True)
    except:
        ## if the previous block of code hasn't changed the symbology it means the patterns of the layers haven't been found
        arcpy.AddWarning("\n*.lyr-patterns of the layers haven't been found.\n  The symbology has been set by default.\n")

    # Create labels
    Lable1 = "X " + str(Xin) + ", Y " + str(Yin)
    Lable2 = "Est.Eror = " + str(R) + " m"           

    # Delete intermediate files
    ## Check to see if intermediate data exist; if they do, then delete them
    intermed = [XY1,XY2,XY3,XY1_UTM,XY2_UTM,XY3_UTM,filepath_UTM,Incenter_UTM]
    for intermed in intermed:
        if arcpy.Exists(intermed):
            arcpy.Delete_management(intermed)

else:
    # Print error message in processing box:
    arcpy.AddError("There is insufficient or wrong data for build of triangle, deriving coordinates\nof observed object and estimation of error of the measurements\n")

