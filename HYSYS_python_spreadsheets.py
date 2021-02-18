'''
                    Aspen HYSYS - Python interface 
                    for spreadsheets connection

Author: Edgar Ivan Sanchez Medina
Email: sanchez@mpi-magdeburg.mpg.de
-------------------------------------------------------------------------------
Date:   21/06/2019
Update: 01/01/2020
'''
import os                          # Import operating system interface
import win32com.client as win32    # Import COM

def Aspen_connection(File_name, Spreadsheet_name, Unit_operation_name,
                     hy_visible=1, active=0):
    '''
    INPUTS
     File_name: String -- File name, including quoting. Example: 'test.hsc'
     Spreadsheet_name: Tuple -- Spreadsheet names in the Aspen HYSYS file
                         including quoting.
     Unit_operation_name: Tuple -- Unit operation name in the Aspen HYSYS 
                         file including quoting.
     hy_visible: Integer -- 1 for Visible, 0 for No Visible
     active: Integer -- 1 for Active, 0 for No Active
     
    OUTPUTS
     Hysys: Class -- It is a class that collects all the important information
                     of the flowsheet.
    NOTE: To check useful pathways to variables in Aspen HYSYS use the Object
            browser from VBA or Matlab and the "HYSYS Customization Guide".
    '''
    
    # 1.0 Obtain the path to the Aspen file
    hyFilePath = os.path.abspath(File_name)
    
    # 2.0 Initialize  Aspen HYSYS application
    print(' ')
    print(' Connecting to the Aspen HYSYS.... Please wait ')
    # This gets the registered name of Aspen HYSYS
    HyApp = win32.Dispatch('HYSYS.Application')
    
    # 3.0 Open Aspen HYSYS file
    if active == 0:
        HyCase = HyApp.SimulationCases.Open(hyFilePath)
    
    # 3.1 Get the active Aspen HYSYS document
    elif active == 1:
        HyCase = HyApp.ActiveDocument
    else:
        raise Exception('Argument for input variable "active" is not valid')
        
    # 4.0 Aspen HYSYS Environment Visible
    HyCase.Visible = hy_visible
    
    # 5.0 Aspen HYSYS File Name
    HySysFile = HyCase.Title.Value
    print(' ')
    print('Aspen HYSYS file name:      ', HySysFile)
    
    # 6.0 Aspen HYSYS Thermodynamic Fluid Package Name
    package_name = HyCase.Flowsheet.FluidPackage.PropertyPackageName
    print('Aspen HYSYS Fluid Package:  ', package_name)
    print(' ')
        
    # 7.0  Spreadsheets
    ss_dict = {}
    for ss in Spreadsheet_name:
        spsh = HyCase.Flowsheet.Operations.Item(ss)
        ss_dict[ss] = spsh
    
    # 8.0 Solver to stop or start the running
    Solver = HyCase.Solver
    
    # 9.0 Material streams
    MaterialStreams = HyCase.Flowsheet.MaterialStreams
    EnergyStreams   = HyCase.Flowsheet.EnergyStreams
    
    # 10.0 Unit operations
    uo_dict = {}
    for uo in Unit_operation_name:
        unop = HyCase.Flowsheet.Operations.Item(uo)
        uo_dict[uo] = unop
    
    # 11.0 Collect everything into a class
    class Hysys:
        pass
    Hysys.HyCase            = HyCase
    Hysys.SS                = ss_dict
    Hysys.Solver            = Solver
    Hysys.MaterialStreams   = MaterialStreams
    Hysys.EnergyStreams     = EnergyStreams
    Hysys.UO                = uo_dict
    
    print(' --- PLEASE! Be aware of the unit handling of this interface--- ')
    print(' --- Python SI Unit Set only --- ')
    print(' ')
    print(' --- It is ALWAYS a good practice to check consistency in units ')
    print('     between your Aspen HYSYS file and the Python interface. --- ')
    print(' ')
    print('**************************************************************** ')
    print(' Python SI unit set: ')
    print('   Temperature:      °C')
    print('   Pressure:         kPa')
    print('   Molar flowrate:   kgmole/s')
    print('   Energy flowrate:  kJ/s')
    print('**************************************************************** ')
    print(' ')
    print( ' Aspen HYSYS-Python Interface has been established succesfully!')
    print(' ')
    return(Hysys)
    