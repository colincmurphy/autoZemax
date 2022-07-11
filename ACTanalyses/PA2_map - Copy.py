import clr, os, winreg
from itertools import islice
import numpy as np
from scipy.optimize import curve_fit
import matplotlib.pyplot as plt
import math
import gc
import sys
from scipy.fft import fft, fftfreq
sys.path.append(r'C:\Users\colinmurphy\Documents\Zemax\ZOS-API Projects\autoZemax') 
import autoACT



# This boilerplate requires the 'pythonnet' module.
# The following instructions are for installing the 'pythonnet' module via pip:
#    1. Ensure you are running Python 3.4, 3.5, 3.6, or 3.7. PythonNET does not work with Python 3.8 yet.
#    2. Install 'pythonnet' from pip via a command prompt (type 'cmd' from the start menu or press Windows + R and type 'cmd' then enter)
#
#        python -m pip install pythonnet

global filePath 
class autoZemax(object):
    class LicenseException(Exception):
        pass
    class ConnectionException(Exception):
        pass
    class InitializationException(Exception):
        pass
    class SystemNotPresentException(Exception):
        pass
    
    global zosFile
    def __init__(self, path=None):
        # determine location of ZOSAPI_NetHelper.dll & add as reference
        aKey = winreg.OpenKey(winreg.ConnectRegistry(None, winreg.HKEY_CURRENT_USER), r"Software\Zemax", 0, winreg.KEY_READ)
        zemaxData = winreg.QueryValueEx(aKey, 'ZemaxRoot')
        NetHelper = os.path.join(os.sep, zemaxData[0], r'ZOS-API\Libraries\ZOSAPI_NetHelper.dll')
        winreg.CloseKey(aKey)
        clr.AddReference(NetHelper)
        import ZOSAPI_NetHelper
        
        # Find the installed version of OpticStudio
        #if len(path) == 0:
        if path is None:
            isInitialized = ZOSAPI_NetHelper.ZOSAPI_Initializer.Initialize()
        else:
            # Note -- uncomment the following line to use a custom initialization path
            isInitialized = ZOSAPI_NetHelper.ZOSAPI_Initializer.Initialize(path)
        
        # determine the ZOS root directory
        if isInitialized:
            dir = ZOSAPI_NetHelper.ZOSAPI_Initializer.GetZemaxDirectory()
        else:
            raise autoZemax.InitializationException("Unable to locate Zemax OpticStudio.  Try using a hard-coded path.")

        # add ZOS-API referencecs
        clr.AddReference(os.path.join(os.sep, dir, "ZOSAPI.dll"))
        clr.AddReference(os.path.join(os.sep, dir, "ZOSAPI_Interfaces.dll"))
        import ZOSAPI

        # create a reference to the API namespace
        self.ZOSAPI = ZOSAPI

        # create a reference to the API namespace
        self.ZOSAPI = ZOSAPI

        # Create the initial connection class
        self.TheConnection = ZOSAPI.ZOSAPI_Connection()

        if self.TheConnection is None:
            raise autoZemax.ConnectionException("Unable to initialize .NET connection to ZOSAPI")

        self.TheApplication = self.TheConnection.CreateNewApplication()
        if self.TheApplication is None:
            raise autoZemax.InitializationException("Unable to acquire ZOSAPI application")

        if self.TheApplication.IsValidLicenseForAPI == False:
            raise autoZemax.LicenseException("License is not valid for ZOSAPI use")

        self.TheSystem = self.TheApplication.PrimarySystem
        if self.TheSystem is None:
            raise autoZemax.SystemNotPresentException("Unable to acquire Primary system")

    def __del__(self):
        if self.TheApplication is not None:
            print('uh oh')
            self.TheApplication.CloseApplication()
            self.TheApplication = None
        
        self.TheConnection = None
    
    def OpenFile(self, filepath, saveIfNeeded):
        if self.TheSystem is None:
            raise autoZemax.SystemNotPresentException("Unable to acquire Primary system")
        self.TheSystem.LoadFile(filepath, saveIfNeeded)

    def CloseFile(self, save):
        if self.TheSystem is None:
            raise autoZemax.SystemNotPresentException("Unable to acquire Primary system")
        self.TheSystem.Close(save)

    def SamplesDir(self):
        if self.TheApplication is None:
            raise autoZemax.InitializationException("Unable to acquire ZOSAPI application")

        return self.TheApplication.SamplesDir

    def ExampleConstants(self):
        if self.TheApplication.LicenseStatus == self.ZOSAPI.LicenseStatusType.PremiumEdition:
            return "Premium"
        elif self.TheApplication.LicenseStatus == self.ZOSAPI.LicenseStatusType.ProfessionalEdition:
            return "Professional"
        elif self.TheApplication.LicenseStatus == self.ZOSAPI.LicenseStatusType.StandardEdition:
            return "Standard"
        elif self.TheApplication.LicenseStatus == self.ZOSAPI.LicenseStatusType.OpticStudioHPCEdition:
            return "HPC"
        else:
            return "Invalid"
    
    def reshape(self, data, x, y, transpose = False):
        """Converts a System.Double[,] to a 2D list for plotting or post processing
        
        Parameters
        ----------
        data      : System.Double[,] data directly from ZOS-API 
        x         : x width of new 2D list [use var.GetLength(0) for dimension]
        y         : y width of new 2D list [use var.GetLength(1) for dimension]
        transpose : transposes data; needed for some multi-dimensional line series data
        
        Returns
        -------
        res       : 2D list; can be directly used with Matplotlib or converted to
                    a numpy array using numpy.asarray(res)
        """
        if type(data) is not list:
            data = list(data)
        var_lst = [y] * x;
        it = iter(data)
        res = [list(islice(it, i)) for i in var_lst]
        if transpose:
            return self.transpose(res);
        return res
    
    def transpose(self, data):
        """Transposes a 2D list (Python3.x or greater).  
        
        Useful for converting mutli-dimensional line series (i.e. FFT PSF)
        
        Parameters
        ----------
        data      : Python native list (if using System.Data[,] object reshape first)    
        
        Returns
        -------
        res       : transposed 2D list
        """
        if type(data) is not list:
            data = list(data)
        return list(map(list, zip(*data)))

if __name__ == '__main__':
    global zos, ZOSAPI, TheApplication, TheSystem, sampleDir
        
    zos = autoZemax()
    
    # load local variables
    ZOSAPI = zos.ZOSAPI
    TheApplication = zos.TheApplication
    TheSystem = zos.TheSystem
    sampleDir = TheApplication.SamplesDir
   
    #fileLocation= os.path.join(os.sep, sampleDir,  r'ACTpol_v28_150GHz_modified.ZOS')
    #TheSystem.LoadFile(fileLocation, False)
    def setup(path):
        # Insert Code Here
        global zosFile
        zosFile = path[:-4]
        print(zosFile)
        TheSystem.LoadFile(path, False)
class configureSystem():
    def clearFields():
        '''
        Delete all existing fields in a ZOS file.

        Returns
        -------
        None.

        '''
        num_fields = TheSystem.SystemData.Fields.NumberOfFields
        print("Clearing fields 1 to " + str(num_fields))
        for i in range(num_fields):
            TheSystem.SystemData.Fields.RemoveField(1)
    def setFields(x, y):
    
        '''
        Clear existing fields. Given two arrays of x and y positions for new fields, 
        set fields in the open ZOS file

        Parameters
        ----------
        x : float
            list of x field values.
        y : flaot
            list of y field values.

        Returns
        -------
        None.

        '''
        configureSystem.clearFields()
        globalNumFields = TheSystem.SystemData.Fields.NumberOfFields
        print("Setting " + str(len(x)) + " new fields")
        for i in range(len(x)):
            TheSystem.SystemData.Fields.AddField(x[i], y[i], 1)
        TheSystem.SystemData.Fields.RemoveField(1)
    def clearWavelengths():
        
    
        '''
        Delete all existing wavelengths in a ZOS file.

        Returns
        -------
        None.

        '''
        num_wavelengths = TheSystem.SystemData.Wavelengths.NumberOfWavelengths
        print("Clearing wavelengths 1 to " + str(num_wavelengths))
        for i in range(num_wavelengths):
            TheSystem.SystemData.Wavelengths.RemoveWavelength(1)
    
    def applyIdealCoating(x):
        '''
        

        Parameters
        ----------
        x : list of integers
            Surfaces to apply ideal coating.

        Returns
        -------
        None.

        '''
        for i in range(len(x)):
            TheSystem.LDE.GetSurfaceAt(x[i]).Coating = "I.0"
    def setWavelengths(x):
    
        '''
        Clear existing wavelengths. Given a list of wavelengths in nanometers,
        set wavelenegths in the open ZOS file

        Parameters
        ----------
        x : float
            list of wavelengths in nm.

        Returns
        -------
        None.

        '''
        configureSystem.clearWavelengths()
        print("Setting " + str(len(x)) + " new fields")
        for i in range(len(x)):
            TheSystem.SystemData.Wavelengths.AddWavelength(x[i], 1)
        TheSystem.SystemData.Wavelengths.RemoveWavelength(1)
class polarizationRotation():
    def makePolarizationPupilMap(surface, wavelength, field, sample, angle):
        '''
        Given the surface, wavelength, field, sample (grid size) and input
        polarization state in terms of angle in degrees from x-axis: compute 
        polarization pupil map, store results in text file, and return the path 
        to the text file

        Parameters
        ----------
        surface : integer
            surface at which OS will calculate polarization pupil map
        wavelength : integer
            tells OS which number wavelength to use when calculating Pupil Map
        field : integer
            tells OS which field number to use when calculating Pupil Map
        sample : integer
            must be between 0 and 40. The grid size of the sampling in the pupil.
        angle : integer
            the angle in degrees from the x axis of the input polarization state

        Returns
        -------
        fileLocation : string
            returns the path to the text file where the output of the pupil map is stored 

        '''
        
        #Convert input angle from degrees to radians
        angle = np.deg2rad(angle)
        MyPolarizationPupilMap = TheSystem.Analyses.New_Analysis(ZOSAPI.Analysis.AnalysisIDM.PolarizationPupilMap)
        ppmSettings = MyPolarizationPupilMap.GetSettings()
        ppmSettings.Save()
        #Settings file must have the same name as the file name
        global zosFile
        configFile = zosFile + r'.CFG'
        ppmSettings.ModifySettings(configFile, 'PPM_WAVE', str(wavelength))
        ppmSettings.ModifySettings(configFile, 'PPM_SURFACE', str(surface))
        ppmSettings.ModifySettings(configFile, 'PPM_FIELD', str(field))
        ppmSettings.ModifySettings(configFile, 'PPM_SAMPLE', str(sample))
        ppmSettings.ModifySettings(configFile, 'PPM_JX', str(np.cos(angle)))
        ppmSettings.ModifySettings(configFile, 'PPM_JY', str(np.sin(angle)))
        
        ppmSettings.LoadFrom(configFile)
        
        MyPolarizationPupilMap.ApplyAndWaitForCompletion()
        ppmResults = MyPolarizationPupilMap.GetResults()
        MyFileName = 'PPM_Results.txt'
        fileLocation = TheApplication.SamplesDir + '\\' + MyFileName
        
        if ppmResults.GetTextFile(TheApplication.SamplesDir + '\\' + MyFileName):
            print('Text file created successfuly \n S: ' + str(surface) + '; W: '
                  + str(wavelength) + '; F: ' + str(field) + '; Sample: ' + str(2*sample + 3) + 
                  '; Jx: ' + str(np.cos(angle)) + '; Jy: ' + str(np.sin(angle)))
        del MyPolarizationPupilMap, ppmResults, ppmSettings
        gc.collect()
        return fileLocation
    def makePolarizationPupilMapAlternate(surface, wavelength, field, sample, Jx, Jy):
        '''
        Given the surface, wavelength, field, sample (grid size) and input Jx, Jy 
        for the polarization state: compute polarization pupil map, store 
        results in text file, and return the path to the text file

        Parameters
        ----------
        surface : integer
            surface at which OS will calculate polarization pupil map
        wavelength : integer
            tells OS which number wavelength to use when calculating Pupil Map
        field : integer
            tells OS which field number to use when calculating Pupil Map
        sample : integer
            must be between 0 and 40. The grid size of the sampling in the pupil.
        Jx : integer
            input polarization Jx
        Jy : TYPE
            input polarization Jy

        Returns
        -------
        fileLocation : string
            returns the path to the text file where the output of the pupil map is stored 

        '''
        MyPolarizationPupilMap = TheSystem.Analyses.New_Analysis(ZOSAPI.Analysis.AnalysisIDM.PolarizationPupilMap)
        ppmSettings = MyPolarizationPupilMap.GetSettings()
        ppmSettings.Save()
        #Settings file must have the same name as the file name
        global zosFile
        configFile = zosFile + r'.CFG'
        ppmSettings.ModifySettings(configFile, 'PPM_WAVE', str(wavelength))
        ppmSettings.ModifySettings(configFile, 'PPM_SURFACE', str(surface))
        ppmSettings.ModifySettings(configFile, 'PPM_FIELD', str(field))
        ppmSettings.ModifySettings(configFile, 'PPM_SAMPLE', str(sample))
        ppmSettings.ModifySettings(configFile, 'PPM_JX', str(Jx))
        ppmSettings.ModifySettings(configFile, 'PPM_JY', str(Jy))
        
        ppmSettings.LoadFrom(configFile)
        
        MyPolarizationPupilMap.ApplyAndWaitForCompletion()
        ppmResults = MyPolarizationPupilMap.GetResults()
        MyFileName = 'PPM_Results.txt'
        fileLocation = TheApplication.SamplesDir + '\\' + MyFileName
        
        if ppmResults.GetTextFile(TheApplication.SamplesDir + '\\' + MyFileName):
            print('Text file created successfuly \n S: ' + str(surface) + '; W: '
                  + str(wavelength) + '; F: ' + str(field) + '; Sample: ' + str(2*sample + 3) + 
                  '; Jx: ' + str(Jx) + '; Jy: ' + str(Jy))
        del MyPolarizationPupilMap, ppmResults, ppmSettings
        gc.collect()
        return fileLocation
    def cleanPolarizationPupilMap(FilePath):
        '''
        Given the path to a text file of a polarization pupil map, organize the
        results into rows. Return the multidimensional array where the values
        are stored

        Parameters
        ----------
        FilePath : string
            Path to the text file where the results of a polarization pupil map
            are stored. 

        Returns
        -------
        Multidimensional array with the values in the polarization pupil map, 
        indexed by in order: Px, Py, Ex, Ey, Intensity, Phase (Deg), Orientation

        '''
        lines = []
        with open(FilePath) as file:
            for line in file:
                lines.append(line)
        lines = lines[36:]
        del lines[1::2]
        #lines = [i.split() for i in lines]
        for i in range(len(lines)):
            lines[i]=lines[i][1::2]
            lines[i]=lines[i].split()
            for j in range(len(lines[i])):
                lines[i][j]=float(lines[i][j])
        lines.pop(-1)
        return lines
    def getAverageOfPupilMap(surface, wavelength, field, sample, angle):
        '''
        Given surface, wavelength, field, sample, angle: create PPM, clean the
        results, average them, and then give back the average rotation of the 
        given field

        Parameters
        ----------
        surface : integer
            surface at which OS will calculate polarization pupil map
        wavelength : integer
            tells OS which number wavelength to use when calculating Pupil Map
        field : integer
            tells OS which field number to use when calculating Pupil Map
        sample : integer
            must be between 0 and 40. The grid size of the sampling in the pupil.
        angle : integer
            the angle in degrees from the x axis of the input polarization state

        Returns
        -------
        float
            average polariztion rotation angle across the pupil for a given field.

        '''
        lines = polarizationRotation.cleanPolarizationPupilMap(polarizationRotation.makePolarizationPupilMap(surface, wavelength, field, sample, angle))
        
        rotations = []
        expected = angle % 180
        
        for row in lines:
            
            if expected == 0 or expected == 180:
                for row in lines:
                    
                    if(row[4] == 0):
                        continue
                    if(row[6] > 160):
                        
                        rotations.append(-np.degrees(np.arctan(row[3]/row[2])))
                        
                    else:
                        
                        rotations.append(np.degrees(np.arctan(row[3]/row[2])))
            else:
            
                for row in lines:
                    
                        if(row[4] == 0):
                            
                            continue
                        rotations.append((np.degrees(np.arctan(row[3]/row[2]))-expected))
                        
        return np.average(rotations)
    
    
if __name__ == '__main__':
    # Your code goes here
    # To begin, input the path to your .ZOS file in the setup function
    
    #setup ACT model (mirrors only)
    path = configFile = os.path.join(os.sep, sampleDir,  r'ACTpol_v28_150GHz_modified.ZOS')
    setup(path)
    
    #Apply ideal coatings
    mirrorSurfaces = autoACT.actInfo.getMirrorSurfaces()
    configureSystem.applyIdealCoating(mirrorSurfaces)
    
    #Apply PA2 fields
    fieldX, fieldY = autoACT.actInfo.getFields(2)
    configureSystem.setFields(fieldX, fieldY)
    
    #Set range of wavelengths
    wavelengths = np.arange(150, 400, 30)
    for i in range(len(wavelengths)):
        wavelengths[i] = 300000/(wavelengths[i])
    configureSystem.setWavelengths(wavelengths)
    
    #Find the maximum and minimum rotation for a variety of wavelengths across 
    #the entire FOV
    averages = []
    
    for j in range(len(fieldX)):
        averages.append(polarizationRotation.getAverageOfPupilMap(10, 1, (j+1), 27, 5))     
        
    #Plot max and min rotations
    for i in range(len(fieldX)):
        fieldX[i]=-fieldX[i]
        fieldY[i]=-fieldY[i]
    plt.scatter(fieldX, fieldY, c = averages)
    plt.colorbar()
    plt.title("Telescope Only Rotation on Sky OpticStudio - PA2")
    
    
    #Cleans up connection to OS. Use when necessary.
    del zos