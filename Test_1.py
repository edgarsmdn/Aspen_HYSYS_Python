"""
                            Test 1 for the                     
                    Aspen HYSYS - Python interface 
                    for spreadsheets connection

Author: Edgar Ivan Sanchez Medina
Email: sanchez@mpi-magdeburg.mpg.de
-------------------------------------------------------------------------------
Date:   05/01/2020

This is a test file for the Aspen HYSYS - Python connection using a flowsheet 
with multiple units.
"""
from HYSYS_python_spreadsheets import aspen_connection
import numpy as np
import matplotlib.pyplot as plt
import time

# 1.0 Data of the Aspen HYSYS file
File = 'Test_1.hsc'
Spreadsheets = ('SS_Flash', 'SS_turbine', 'SS_Distillation')
Units = ('Cooler', 'Flash Drum', 'Heater', 'Valve', 'Reactor', 'Distillation Column', 'Turbine', 'Pump')

# 2.0 Perform connection
Test_1 = aspen_connection(File, Spreadsheets, Units)
Turbine = Test_1.spreadsheets['SS_turbine']
Efficiency = Turbine.Cell(1, 0)            # .Cell(Column,Row) starting from 0
Generation = Turbine.Cell(1, 1)
ori_eff = Efficiency.CellValue
solver = Test_1.solver

print('Original Turbine efficiency: ', ori_eff)

# 3.0 CONVERGENCE TIME PROBLEM
points = 10
eff = np.zeros(points)
gen = np.zeros(points)
fig, axs = plt.subplots(1, 2)

# 3.1 Not waiting for solver to stop
for i in range(points):
    eff[i] = Efficiency.CellValue
    gen[i] = Generation.CellValue
    Efficiency.CellValue = eff[i] + 1

# 3.11 Plot
axs[0].plot(eff, gen, 'ro')
axs[0].set_xlabel('Turbine Efficiency')
axs[0].set_ylabel('Energy Generation [kJ/s]')
axs[0].set_title('Not waiting Solver')

# 3.2 Come back to original
eff = np.zeros(points)
gen = np.zeros(points)
Efficiency.CellValue = ori_eff

# 3.3 Waiting for solver to stop
for i in range(points):
    eff[i] = Efficiency.CellValue
    gen[i] = Generation.CellValue
    solver.CanSolve = False                  # Turn off the solving mode
    Efficiency.CellValue = eff[i] + 1
    solver.CanSolve = True                   # Turn on the solving mode
    while solver.IsSolving:
        time.sleep(0.001)

# 3.31 Plot
axs[1].plot(eff, gen, 'ro')
axs[1].set_xlabel('Turbine Efficiency')
axs[0].set_ylabel('Energy Generation [kJ/s]')
axs[1].set_title('Waiting solver')
plt.tight_layout()
plt.show()
plt.close()


# Return to original
Efficiency.CellValue = ori_eff
