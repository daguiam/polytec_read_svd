import numpy as np


import win32com.client
import os
import json


class ClassData:
    pass


from json import JSONEncoder

class NumpyArrayEncoder(JSONEncoder):
    FORMAT_SPEC = '@@{}@@'

    def default(self, obj):
        if isinstance(obj, np.ndarray):
            return obj.tolist()
        
#         if isinstance(obj, np.ndarray):
#             return self.FORMAT_SPEC.format(id(obj))
# #             return obj.tolist()
#         return JSONEncoder.default(self, obj)


def GetPointData(filename, domainname, channelname, signalname, displayname, point, frame):
    """
    % [x,y,usd] = GetPointData(filename, domainname, channelname, signalname, displayname,
    %   point, frame)
    % 
    % Gets original or user defined data from a polytec file.
    %
    % filename is the path of the .pvd or .svd file
    % domainname is the name of the domain, e.g. 'FFT' or 'Time'
    % channelname is the name of the channel, e.g. 'Vib' or 'Ref1' or 'Vib &
    %   Ref1' or 'Vib X' or 'Vib Y' or 'Vib Z'.
    % signalname is the name of the signal, e.g. 'Velocity' or 'Displacement'
    % displayname is the name of the display, e.g. 'Real' or 'Magnitude' or
    %   'Samples'. If the display name is 'Real & Imag.' the data is returned
    %   as complex values.
    % point is the (1-based) index of the point to get data from. If point is
    %   0 the data of all points will be returned. y will contain the data of
    %   point i at row index i.
    % frame is the frame number of the data. for data acquired in MultiFrame
    %   mode, 0 is the averaged frame and 1-n are the other frames. For user
    %   defined datasets the frame number is in the range 1-n where n is the
    %   number of frames in the user defined dataset. For all other data,
    %   use frame number 0.
    %
    % returns x, the x axis values of the data
    % returns y, the data. colomns correspond to the x-axis, rows to the point
    %   index. for point = 0: rows for points that have no data are set to zeros.
    % returns usd, a struct describing the signal
    """
    
    file = win32com.client.Dispatch('PolyFile.PolyFile') 
    #Make sure that you can write to the file
    
    file_path = filename
#     file_path = os.path.join(".", filename)

    file.ReadOnly = False
    file.Open(file_path)
    
    usd = ClassData


    pointdomains = file.GetPointDomains();
    pointdomain = pointdomains.Item(domainname);
    channel = pointdomain.Channels.Item(channelname);
    signal = channel.Signals.Item(signalname);
    display = signal.Displays.Item(displayname);

    signalDesc = signal.Description;
    xaxis = signalDesc.XAxis;
    yaxis = signalDesc.YAxis;

    x = np.linspace(xaxis.Min, xaxis.Max, xaxis.MaxCount)
    usd.Name = signalDesc.Name;
    usd.Complex = signalDesc.Complex;
    usd.DataType = signalDesc.DataType;
    usd.DomainType = signalDesc.DomainType;
    usd.FunctionType = signalDesc.FunctionType;
    usd.PowerSignal = signalDesc.PowerSignal;
    usd.Is3D = (signalDesc.ResponseDOFs.Count > 0) and (not str(signalDesc.ResponseDOFs.Direction).find('ptcVector'))
    responseDOFs = signalDesc.ResponseDOFs;

    usd.ResponseDOFs = []

    if responseDOFs.Count == 0:
        usd.ResponseDOFs = []
    else:
        for i in range(1,responseDOFs.Count+1):
            usd.ResponseDOFs.append(responseDOFs.Item(i))

    referenceDOFs = signalDesc.ReferenceDOFs
    usd.ReferenceDOFs = []
    if referenceDOFs.Count == 0:
        usd.ReferenceDOFs = []
    else:
        for i in range(1,referenceDOFs.Count):
            usd.ReferenceDOFs.append(referenceDOFs.Item(i))

    usd.DbReference = signalDesc.DbReference
    usd.XName = xaxis.Name
    usd.XUnit = xaxis.Unit
    usd.XMin = xaxis.Min
    usd.XMax = xaxis.Max
    usd.XCount = xaxis.MaxCount
    usd.YName = yaxis.Name
    usd.YUnit = yaxis.Unit
    usd.YMin = yaxis.Min
    usd.YMax = yaxis.Max

    datapoints = pointdomain.datapoints

    if (point == 0):
    #     % get data of all points
        y = []

        nr_datapoints = datapoints.count

#         y = np.zeros((int(nr_datapoints),int(usd.XCount)))
        for i in range(nr_datapoints):
            datapoint = datapoints.Item(i+1);

            ytemp = np.array(datapoint.GetData(display, frame));
#             y[i,:] = ytemp
            y.append(ytemp)
    file.Close()
    
    return (x,y,usd)


def GetXYZCoordinates(filename, point):
    """
    % Gets XYZ coordinates of the scan points from a polytec file.
    %
    % This is only possible for files containing 3D geometry or that have
    %   a distance to the object specified. Otherwise there will be an
    %   error message.
    %
    % filename is the path of the .svd file
    % 
    % point is the (1-based) index of the point to get the coordinates from. If point is
    %   0 the coordiantes of all points will be returned. XYZ will contain the data of
    %   point i at row index i.
    %
    % returns the xyz coordinates. columns correspond to the geometry X, Y, Z
    %   in meter, rows to the point index.
    """
    file = win32com.client.Dispatch('PolyFile.PolyFile') 
    #Make sure that you can write to the file
    file.ReadOnly = False
    
    file_path = filename
    file.Open(file_path)


    measpoints = file.Infos.MeasPoints;
 
    if point == 0:
        XYZ = []
        for i in range(measpoints.count):
            measpoint=measpoints.Item(np.int32(i+1));
            [X,Y,Z] = measpoint.CoordXYZ();
#             XYZ(i,:)=[X,Y,Z];                  
            XYZ.append([X,Y,Z])
    return XYZ




def CreateDataDict(filename, domainname, vib_channelname, vib_signalname, ref_channelname, ref_signalname, displayname, point, frame):

    print("Opening", filename)
    time,displacement,usd =  GetPointData(filename, domainname, vib_channelname, vib_signalname, displayname, point, frame)
    time,voltage,usd =  GetPointData(filename, domainname, ref_channelname, ref_signalname, displayname, point, frame)
    
    XYZ = GetXYZCoordinates(filename, point)
    nr_points = len(displacement)
    print("Loaded %d measurements points with %d samples"%(nr_points, len(displacement[0])))
    data = dict()

    data["filename"] = filename
    data["domainname"] = domainname
    data["vib_channelname"] = vib_channelname
    data["vib_signalname"] = vib_signalname
    data["ref_channelname"] = ref_channelname
    data["ref_signalname"] = ref_signalname
    data["displayname"] = ref_signalname
    data["nr_points"] = nr_points

    for i in range(nr_points):
        data[i] = dict()
        data[i]['x'] = XYZ[i][0]
        data[i]['y'] = XYZ[i][1]
        data[i]['z'] = XYZ[i][2]
        data[i]['time'] = time
        data[i]['voltage'] = voltage[i]
        data[i]['displacement'] = displacement[i]
        
    out_filename = filename+".json"
    print("Saving data to %s"%(out_filename))
    
    with open(out_filename, "w") as outfile:
        json.dump(data, outfile,  cls=NumpyArrayEncoder, indent=4)
    print("Done saving data to %s"%(out_filename))
    return data



def save_dict2json(data, out_filename):

    
    with open(out_filename, "w") as outfile:
        json.dump(data, outfile,  cls=NumpyArrayEncoder, indent=4)
        


if __name__ == "__main__":


            
    filename = "SA_01_left.svd"

    domainname = "Time"
    vib_channelname = "Vib"
    vib_signalname = "Displacement"
    ref_channelname = "Ref1"
    ref_signalname = "Voltage"
    displayname = "Samples"
    point = 0
    frame = 0

    data = CreateDataDict(filename, domainname, vib_channelname, vib_signalname, ref_channelname, ref_signalname, displayname, point, frame)