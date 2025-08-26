#ASPEN Simulation: VLE flash of ethanol, water and benzene

#Import libraries
import os                         # Import operating system interface
import win32com.client as win32   # Import COM

#Specify file name/Path
file = r"\\nask.man.ac.uk\home$\Documents\SOE Summer Internship\Trial\VLE.bkp"
aspen_Path = os.path.abspath(file)

#Launch ASPEN PLus
aspen = win32.Dispatch('Apwn.Document') #Launch ASPEN Plus
aspen.InitFromArchive2(aspen_Path) #Load simulation

#Reinitialise simulation
aspen.Engine.Reinit()  

#Input data
#Feed composition
#z_ETOH
aspen.Tree.FindNode(r"\Data\Streams\S1\Input\FLOW\MIXED\ETHANOL").Value = 0.8
#z_H2O
aspen.Tree.FindNode(r"\Data\Streams\S1\Input\FLOW\MIXED\WATER").Value = 0.1
#z_BENZENE
aspen.Tree.FindNode(r"\Data\Streams\S1\Input\FLOW\MIXED\BENZENE").Value = 0.1
#Pressure
aspen.Tree.FindNode(r"\Data\Blocks\VLE\Input\PRES").Value = 1 #atm

#Re-run the simulation
aspen.Engine.Run2() #for COM interface

#Output data
#VLE data
#Liquid composition 
x_ETOH = aspen.Tree.FindNode(r"\Data\Blocks\VLE\Output\X\ETHANOL").Value
x_H2O = aspen.Tree.FindNode(r"\Data\Blocks\VLE\Output\X\WATER").Value
x_BENZENE = aspen.Tree.FindNode(r"\Data\Blocks\VLE\Output\X\BENZENE").Value 
#Vapour composition
y_ETOH = aspen.Tree.FindNode(r"\Data\Blocks\VLE\Output\Y\ETHANOL").Value
y_H2O = aspen.Tree.FindNode(r"\Data\Blocks\VLE\Output\Y\WATER").Value
y_BENZENE = aspen.Tree.FindNode(r"\Data\Blocks\VLE\Output\Y\BENZENE").Value

#Molar enthalpy 
#Feed
S1_H = aspen.Tree.FindNode(r"\Data\Streams\S1\Output\HMX\MIXED").Value*4.184
#Vapour product
S2_H = aspen.Tree.FindNode(r"\Data\Streams\S2\Output\HMX\MIXED").Value*4.184
#Liquid product
S3_H = aspen.Tree.FindNode(r"\Data\Streams\S3\Output\HMX\MIXED").Value*4.184

#Report results
print('x_ETOH =', x_ETOH)
print('x_H2O =', x_H2O)
print('x_BENZENE =', x_BENZENE)

print('y_ETOH =', y_ETOH)
print('y_H2O =', y_H2O)
print('y_BENZENE =', y_BENZENE)

print('Feed Enthalpy =', S1_H, 'kJ/kmol')
print('Vapour Product Enthalpy =', S2_H, 'kJ/kmol')
print('Liquid Product Enthalpy =', S3_H, 'kJ/kmol')