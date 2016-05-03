#!/usr/bin/python

"""
Title: cases2kml.py 

Author: Chris Jewell <chris.jewell@warwick.ac.uk> (c) 2010

License: GNU Public License v3 <http://www.gnu.org/licenses/gpl.html>

Purpose: Turns a specified CSV format for disease case
           into KML for plotting in Google Earth

Details:

This package provides both a class and a wrapper function to
convert CSV cases files to KML.  The required CSV format is:

   <long>,<lat>,<report date>[,<id>]

Note: the <id> field above is optional

Although the class Cases2kml is provided for use in software,
it is recommended that the function cases2kml is used for one
off conversions.  eg:
  
   >>> from cases2kml import cases2kml
   >>> cases2kml("myInput.csv","myOutput.kmz",'M',3.0)

"""

import os,sys
import csv
from optparse import OptionParser
from time import strptime
from datetime import date,timedelta
from math import sqrt
from zipfile import ZipFile,ZIP_DEFLATED
from cStringIO import StringIO

version = "1.0-6beta"


def _addMonths(startTime,numMonths):
    year = startTime.year + numMonths // 12
    month = startTime.month + numMonths % 12
    if month > 12:
        month = month - 12
        year += 1
    try:
        return startTime.replace(year=year,month=month,day=1)
    except ValueError as err:
        print "Start time: %s" % startTime.isoformat()
        print "Num months: %i" % numMonths
        print "Month = %i" % month
        raise err

class DateError(Exception):
    pass

class _Meshblock:

    def __init__(self,x,y,id):
        self.x = x
        self.y = y
        self.id = id
        self.cases = {}

    def addCase(self,date):
        if date in self.cases:
            self.cases[date] += 1
        else:
            self.cases[date] = 1
            
    def aggregate(self,aggrDates):
        pass

    def maxCases(self):
        """Returns the maximum number of cases in any date aggregation"""
        maxNum = 0
        for key,val in self.cases.iteritems():
            if val > maxNum:
                maxNum = val
        return maxNum


class Cases2kml:

    def __init__(self,aggrUnit,aggrCount,pointMag,colour):
        self.aggrUnit = aggrUnit
        self.aggrCount = aggrCount
        self.pointMag = float(pointMag)
        self.colour = colour
        
    def __placemark(self,x,y,id,date,numCases):
        """Serializes a placemark"""
        
        pmString = StringIO()

        if self.aggrUnit == 'D':
            end = date + timedelta(days=self.aggrCount)
            
        elif self.aggrUnit == 'M':
            end = _addMonths(date,self.aggrCount)

        elif self.aggrUnit == 'Y':
            end = date.replace(year=date.year+self.aggrCount)
            
        else:
            raise ValueError("Invalid aggregation unit!")
        

        description = """Period beginning: %s
<br>Period ending: %s
<br>Longitude: %f
<br>Latitude: %f
<br>Number of cases: %i""" % (date.isoformat(), end.isoformat(), x, y, numCases)

        pointSize = sqrt(float(numCases))
        
        pmString.write( "   <Placemark>\n" )
        pmString.write( "    <description><![CDATA[" + description + "]]></description>\n" )
        pmString.write( "    <TimeSpan>\n" )
        pmString.write( "     <begin>" + date.isoformat() + "</begin>\n" )
        pmString.write( "     <end>" + end.isoformat() + "</end>\n" )
        pmString.write( "    </TimeSpan>\n" )
        pmString.write( "    <Point>\n" )
        pmString.write( "     <coordinates>" + str(x) + "," + str(y) + ",0</coordinates>\n" )
        pmString.write( "    </Point>\n" )
        pmString.write( "    <Style>\n" )
        pmString.write( "     <IconStyle>\n" )
        pmString.write( "      <Icon>http://maps.google.com/mapfiles/kml/shapes/shaded_dot.png</Icon>\n" )
        pmString.write( "      <color>" + self.colour + "</color>\n" )
        pmString.write( "      <scale>" + str(pointSize*self.pointMag) + "</scale>\n" )
        pmString.write( "     </IconStyle>\n" )
        pmString.write( "    </Style>\n" )
        pmString.write( "  </Placemark>\n" )

        return pmString.getvalue()



    def readCSV(self,inputfile, fieldMap, datefmt="%Y-%m-%d", quoteChar="", delimChar=",",progressUpdateFunc=None):
        """csv2kml converts a CSV file of event times into a time-aggregated
    KML file. Command line options are:
        inputFile: CSV file
        fielaMap: a dictionary of x,y,date (keys) and col number
        aggr: the aggregation level [D,M,Y].
        datefmt: the Python time.strftime() format code.
        """

        self.meshblocks = {}
        
        xField = fieldMap['x']
        yField = fieldMap['y']
        dateField = fieldMap['date']
        self.maxDate = date.min
        self.minDate = date.max
        
        #Quoting
        isQuoted = csv.QUOTE_NONE
        if quoteChar != "":
            isQuoted = csv.QUOTE_MINIMAL

        # Open CSV file
        csvFile = open(inputfile,"rb")
        
        # First pass - get max and min dates
        reader = csv.reader(csvFile,quoting=isQuoted,quotechar=str(quoteChar),delimiter=str(delimChar))
        reader.next() # Skip header
        progressUpdateFunc("Inspecting file....")
        
        counter = 0
        for row in reader:
            if progressUpdateFunc != None and (counter % 1000) == 0:
                progressUpdateFunc()
                
            counter += 1
            
            try:
                parseDate = strptime(row[dateField],datefmt)
            except ValueError as err:
                raise DateError(err.args)
            
            myDate = date(*parseDate[0:3])
                
            self.maxDate = max(myDate,self.maxDate)
            self.minDate = min(myDate,self.minDate)
            
        # Reset file
        csvFile.seek(0)
        
        # Second pass - get the data in
        reader = csv.reader(csvFile,quoting=isQuoted,quotechar=str(quoteChar),delimiter=str(delimChar))
        reader.next() # Skip header
        progressUpdateFunc("Reading data....")
        for row in reader:
            if progressUpdateFunc != None and (counter % 1000) == 0:
                progressUpdateFunc()
            
            counter += 1

            x = float(row[xField])
            y = float(row[yField])
            
            try:
                parseDate = strptime(row[dateField],datefmt)
            except ValueError as err:
                raise DateError(err.args)
            
            myDate = date(*parseDate[0:3])

            id = None
            if len(row) > 3:
                id = row[3]

            # Set date to beginning of aggregation period
            if self.aggrUnit == 'D':
                numDays = (myDate - self.minDate).days
                numAggUnits = numDays // self.aggrCount
                myDate = self.minDate + timedelta(days=self.aggrCount * numAggUnits)
                
            elif self.aggrUnit == 'M':
                diffMonths = (myDate.year - self.minDate.year)*12 + myDate.month - self.minDate.month
                numAggUnits = diffMonths // self.aggrCount
                myDate = _addMonths(self.minDate,numAggUnits*self.aggrCount)
                
            elif self.aggrUnit == 'Y':
                diffYears = myDate.year - self.minDate.year
                numAggUnits = diffYears // self.aggrCount
                year = self.minDate.year + numAggUnits * self.aggrCount
                myDate = myDate.replace(year=year,month=1,day=1)
                
            else:
                raise ValueError("Invalid aggregation unit!")

            lockey = row[xField]+row[yField]
                
            if lockey in self.meshblocks:
                self.meshblocks[lockey].addCase(myDate)
            else:
                self.meshblocks[lockey] = _Meshblock(x,y,id)
                self.meshblocks[lockey].addCase(myDate)

        # Get max cases number
        self.maxNum = 0
        for key,meshblock in self.meshblocks.iteritems():
            myMaxCases = meshblock.maxCases()
            if myMaxCases > self.maxNum:
                self.maxNum = myMaxCases


    def serialize(self,docName,progressFunction=None):
        # Now serialise the meshblocks into kml        
        numMeshBlocks = len(self.meshblocks)
        serialized = StringIO()
        
        # Write kml header
        serialized.write("""<?xml version="1.0" encoding="us-ascii"?>
<kml xmlns="http://earth.google.com/kml/2.1">
<Document>\n <name>""" + docName + "</name>\n")

        counter = 0
        # Loop through meshblocks and serialize
        for key,meshblock in self.meshblocks.iteritems():
            
            serialized.write( " <Folder>\n" )
            serialized.write( "  <name>" + meshblock.id + "</name>\n" )

            for date,numCases in meshblock.cases.iteritems():

                serialized.write( self.__placemark(meshblock.x,meshblock.y,meshblock.id,date,numCases) + "\n" )

            serialized.write( " </Folder>\n" )
            
            if progressFunction != None:
                progressFunction(float(counter)/numMeshBlocks * 100)
                
            counter += 1

        # Write footer
        serialized.write( "</Document>\n" )
        serialized.write( "</kml>" )
        
        return serialized.getvalue()


def cases2kml(inputfile,outputfile,aggr='M',mag=1.0,colour='FF0000FF',dateformat='%Y-%m-%d'):
    """cases2kml takes a CSV file of cases, and outputs a file in KMZ format for viewing in Google Earth.

Arguments:
  inputfile - the input CSV file
  outputfile - the name of the .kmz file to write to
  aggr - the aggregation level, currently supported values are D, M, Y
  mag - the magnification level for the map points in Google Earth
  dateformat - the strftime formatting code for time

Details:
  The format of the CSV file must conform to the fields:

      <longitude>,<latitude>,<case date>[,<id>]

  The longitude and latitude fields must be supplied in WGS 84 geographic coordinates (ie decimal long and lat).  The optional id field is for your administrative purposes only.
"""
    
    
    print "Converting '" + inputfile + "'"
    print "Aggregation level: '" + aggr + "'"

    sys.stdout.flush()

    kmlWriter = Cases2kml(aggr,mag,1,colour)
    
    fieldMap = {'x': 0, 'y': 1, 'date': 2}

    kmlWriter.readCSV(inputfile,fieldMap,dateformat)

    # Serialize to kmz file
    kmz = ZipFile(outputfile,"w",ZIP_DEFLATED)
    kmz.writestr("doc.kml",kmlWriter.serialize(outputfile))
    kmz.close()

    print "Done\n"



if __name__ == "__main__":
    
    # Command line options

    usage = """usage: %prog [options] <input csv> <output kml>"""
    optparse = OptionParser(usage=usage,version="%prog "+version)
    
    optparse.add_option("-a", "--aggregate", dest="aggregate",
                        action="store",default="M",
                        help="set level of time aggregation [D,M,Y] [default: %default]")
    optparse.add_option("-d", "--date-format", dest="dateformat",
                        default="%Y-%m-%d",
                        help="Python time.strftime() format representing date [default: %default]")
    optparse.add_option("-m", "--magnify", dest="mag",
                        default=1.0,
                        help="Magnification factor for plotted points [default: %default]")
    optparse.add_option("-c","--colour", dest="col",
                        default='FF0000FF',
                        help="Map point colour expressed as AABBGGRR (ie KML spec)")

    (options,args) = optparse.parse_args()

    # Check args
    if len(args) != 2:
        optparse.print_help()
        sys.exit(1)

    # Check options
    if options.aggregate not in ['D','M','Y']:
        print "Unrecognised aggregation level: '" + options.aggregate + "'"
        sys.exit(1)
 
    cases2kml(args[0],args[1],options.aggregate,options.mag,options.col,options.dateformat)

    sys.exit(0)
