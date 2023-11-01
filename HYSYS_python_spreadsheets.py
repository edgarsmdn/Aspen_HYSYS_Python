"""
                    Aspen HYSYS - Python interface 
                    for spreadsheets connection

Author: Edgar Ivan Sanchez Medina
Email: sanchez@mpi-magdeburg.mpg.de
-------------------------------------------------------------------------------
Date:   21/06/2019
Update: 01/01/2020
"""
# built-in
import os                          # Import operating system interface
from typing import Tuple

# third-party
import win32com.client as win32    # Import COM


class Hysys:
    hysys_case = ...
    solver = ...
    spreadsheets = ...
    material_streams = ...
    energy_streams = ...
    unit_operation = ...


def aspen_connection(file_name: str,
                     spreadsheet_name: Tuple[str, ...],
                     unit_operation_name: Tuple[str, ...],
                     hy_visible: int = 1,
                     active: int = 0):
    """
    method for connection to hysys (.hsc) files
    Args:
        file_name: File name, including quoting. Example: 'test.hsc'
        spreadsheet_name: Spreadsheet names in the Aspen HYSYS file including quoting.
        unit_operation_name: Unit operation name in the Aspen HYSYS file including quoting.
        hy_visible: 1 for Visible, 0 for No Visible
        active: 1 for Active, 0 for No Active

    Returns:
        Hysys: Class -- It is a class that collects all the important information of the flowsheet.

    NOTE: To check useful pathways to variables in Aspen HYSYS use the Object
            browser from VBA or Matlab and the "HYSYS Customization Guide".
    """

    # 1.0 Obtain the path to the Aspen file
    absolute_path = os.path.abspath(file_name)
    
    # 2.0 Initialize  Aspen HYSYS application
    print(' ')
    print(' Connecting to the Aspen HYSYS.... Please wait ')

    # This gets the registered name of Aspen HYSYS
    # ToDo read dock of win32.Dispatch
    hysys_app = win32.Dispatch('HYSYS.Application')
    
    # 3.0 Open Aspen HYSYS file
    if active == 0:
        hysys_case = hysys_app.SimulationCases.Open(absolute_path)
    
    # 3.1 Get the active Aspen HYSYS document
    elif active == 1:
        hysys_case = hysys_app.ActiveDocument
    else:
        raise Exception('Argument for input variable "active" is not valid')
        
    # 4.0 Aspen HYSYS Environment Visible
    hysys_case.Visible = hy_visible
    
    # 5.0 Aspen HYSYS File Name
    file_name = hysys_case.Title.Value
    print(' ')
    print('Aspen HYSYS file name:      ', file_name)
    
    # 6.0 Aspen HYSYS Thermodynamic Fluid Package Name
    package_name = hysys_case.Flowsheet.FluidPackage.PropertyPackageName
    print('Aspen HYSYS Fluid Package:  ', package_name)
    print(' ')
        
    # 7.0  Spreadsheets
    ss_dict = {}
    for ss in spreadsheet_name:
        spsh = hysys_case.Flowsheet.Operations.Item(ss)
        ss_dict[ss] = spsh
    
    # 8.0 Solver to stop or start the running
    solver = hysys_case.Solver
    
    # 9.0 Material streams
    material_streams = hysys_case.Flowsheet.MaterialStreams
    energy_streams = hysys_case.Flowsheet.EnergyStreams
    
    # 10.0 Unit operations
    uo_dict = {}
    for uo in unit_operation_name:
        unop = hysys_case.Flowsheet.Operations.Item(uo)
        uo_dict[uo] = unop
    
    # 11.0 Collect everything into a class
    Hysys.hysys_case = hysys_case
    Hysys.spreadsheets = ss_dict
    Hysys.solver = solver
    Hysys.material_streams = material_streams
    Hysys.energy_streams = energy_streams
    Hysys.unit_operation = uo_dict
    
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
    print(' Aspen HYSYS-Python Interface has been established succesfully!')
    print(' ')
    return Hysys
    