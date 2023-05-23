# Aspen HYSYS-Python connection using spreadsheets

This repo contains an Aspen HYSYS - Python connection using spreadsheets.

Since accesing the variables paths in Aspen HYSYS is sometimes problematic. The use of spreadsheets is an easy and fast way to access the variables we want.

<img align="center" src="https://github.com/edgarsmdn/Aspen_HYSYS_Python/blob/main/figures/HYSYS_Python.PNG" width="1000">


## Function ```Aspen_connection``` in ```HYSYS_python_spreadsheets.py```

```
Aspen_connection(File_name, Spreadsheet_name, Unit_operation_name, hy_visible=1, active=0)
```

#### Inputs

&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 1. ```File_name``` of the Aspen HYSYS file you are working with.

&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; E.g. ``` 'Test_1.hsc' ```

&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 2. ```Spreadsheet_name``` is a list of names for the specific spreadsheets within the Aspen HYSYS file that we are connecting with Python.

&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; E.g. ``` ('SS_Flash', 'SS_turbine', 'SS_Distillation')```

&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 3. ```Unit_operation_name``` is a list of the names of the unit operations present within the Aspen HYSYS file. This is useful for example when dealing with distillation columns and their specific flowsheet window.
                
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; E.g. ``` ('Cooler', 'Flash Drum', 'Heater', 'Valve', 'Reactor', 
                'Distillation Column', 'Turbine', 'Pump')```
                
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 4. ```hy_visible``` whether to make Aspen HYSYS visible or not. 1=visible, 0=No Visible. Default 1.

&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 5. ```active``` whether the Aspen HYSYS file is currently active or not. 1=Active, 0=No Active. Default 0.

#### Output

The output is a class ```Hysys```   with the following methods:

&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 1. ```Hysys.HyCase``` is the complete Aspen HYSYS case. From here you can access all the variables usign the specific variable paths. 
To find this paths you can use the "Object browser" from Excel VBA or Matlab and the "HYSYS Customization Guide" pdf file. As I said, this is sometimes problematic. So you can work directly on spreadsheets. 

&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 2. ```Hysys.SS``` is a dictionary with the connections to the spreadsheets.

&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 3. ```Hysys.Solver``` is the solver in Aspen HYSYS. This is specially useful to turn it ON/OFF when changing input values to the simulation.

&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 4. ```Hysys.MaterialStreams``` is the connection to material streams in case you want to use (and know) the full path to the variable of interest here.

&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 5. ```Hysys.EnergyStreams``` is the connection to energy streams in case you want to use (and know) the full path to the variable of interest here.

&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 6. ```Hysys.UO``` is a dictionary with the connection to the unit operations. With this you can access the specific options of each unit if you know the path.

## Example case-study ```Test_1.py```

An example of use is provided in the ```Test_1.py``` where I show the importance of waiting for the solver to finish when working with the Python connection. Enjoy!

## License

This repository contains a [MIT LICENSE](https://github.com/AlxndrKlbk/pyhys/blob/main/LICENSE)
