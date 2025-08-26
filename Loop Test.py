#Sensitivity loop test: Study impact of varying pressure on vapour composition
#ASPEN Simulation: VLE flash of ethanol, water and benzene

#Import libraries
import os                         # Import operating system interface
import win32com.client as win32   # Import COM
import numpy as np
import matplotlib.pyplot as plt

#Specify file name/Path
file = r"\\nask.man.ac.uk\home$\Documents\SOE Summer Internship\Trial\VLE.bkp"
aspen_Path = os.path.abspath(file)

#Launch ASPEN PLus
aspen = win32.Dispatch('Apwn.Document') #Launch ASPEN Plus
aspen.InitFromArchive2(aspen_Path) #Load simulation

#Reinitialise simulation
aspen.Engine.Reinit()  

#Input data
#Pressure
P_range = np.linspace(0.1, 1, 10) #atm

#Results array
y_ETOH_array = []
y_H2O_array = []
y_BENZENE_array = []

for P in P_range:
    aspen.Engine.Reinit() #Reinitialise simulation
    aspen.Tree.FindNode(r"\Data\Blocks\VLE\Input\PRES").Value = P 
    aspen.Engine.Run2() #run simulation

    #Read compositions
    y_ETOH = aspen.Tree.FindNode(r"\Data\Blocks\VLE\Output\Y\ETHANOL").Value
    y_H2O = aspen.Tree.FindNode(r"\Data\Blocks\VLE\Output\Y\WATER").Value
    y_BENZENE = aspen.Tree.FindNode(r"\Data\Blocks\VLE\Output\Y\BENZENE").Value 

    #Store output data in results array
    y_ETOH_array.append(y_ETOH)
    y_H2O_array.append(y_H2O)
    y_BENZENE_array.append(y_BENZENE)

#Print results
print(P_range)
print(y_ETOH_array)
print(y_H2O_array)
print(y_BENZENE_array)

#Plot results
#Liquid composition
plt.plot(P_range, y_ETOH_array, 'r', label='y_ETOH')
plt.plot(P_range, y_H2O_array, 'b', label='y_H2O')
plt.plot(P_range, y_BENZENE_array, 'g', label='y_BENZENE')

plt.xlabel('Pressure (atm)')
plt.ylabel('Vapour composition')

plt.legend()
plt.grid(True)
plt.tight_layout()
plt.show()





