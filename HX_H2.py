#Stream data extraction for H2 chemical looping process  

#Import libraries
import os                         # Import operating system interface
import win32com.client as win32   # Import COM
import pandas as pd               # Export data to Excel
from openpyxl import load_workbook  # To present data on multiple sheets on Excel

#Specify file name/path 
file = r"\\nask.man.ac.uk\home$\Documents\SOE Summer Internship\Trial\MERHEN.bkp"
aspen_Path = os.path.abspath(file)

print('Loading Simulation...')
#Launch ASPEN Plus 
aspen = win32.Dispatch('Apwn.Document') 
#Load simulation
aspen.InitFromArchive2(aspen_Path) 
#Reinitialise simulation
aspen.Engine.Reinit()  
#Run simulation 
aspen.Engine.Run2() 
print('Simulation Loaded Successfully')

#Part 1: Data Extraction 
print('Extracting Data...')
#Identify all heat exchanger blocks 
#Access all blocks in the simulation 
blocks = aspen.Tree.Elements("Data").Elements("Blocks")

#Main parameter list
HX = []
T_in = []
T_out = []
Type =[]
Q = []

#Iterate through each block 
for block in blocks.Elements:  
    block_name = block.Name  #Obtain the names of each block 
    #Identify HX based on block type 
    if block_name.startswith(('E', 'K', 'F')):  #Heat exchanger blocks are identified based on naming system 
        HX.append(block_name)  #Add only heat exchanger blocks to HX list

#Extract stream data for each heat exchanger block 
for block_name in HX:
        
    #Pathway to calculated heat duty 
    QCALC = aspen.Tree.Elements("Data").Elements("Blocks").Elements(block_name).Elements("Output").Elements("QCALC").Value
    #Convert string to float
    if QCALC is not None:
        QCALC = float(QCALC)
    else:
        QCALC = 0

    #Intermediate list to store input and output temperatures
    temperature_data = []
    #Pathway to stream results table 
    Table = aspen.Tree.Elements("Data").Elements("Blocks").Elements(block_name).Elements("Stream Results").Elements("Table")
    #Iterate to identify temperature elements under "Table" list 
    for element in Table.Elements:
        #Identify the temperature elements in the resuls table of each heat exchanger block 
        if "Temperature" in element.Name: 
            #Extract stream temperature
            stream_temperature = element.Value 
            #Convert string to float
            stream_temperature = float(stream_temperature)
            #Store extracted temperatures in stream_data list, 1st column is input and 2nd column is output 
            temperature_data.append(stream_temperature)
            #Break the loop once temperatures of both inlet and outlet stream are extracted, 3rd element is an empty node
            if len(temperature_data) == 2:
                break
    
    #Determine stream type 
    input_temp = temperature_data[0]
    output_temp = temperature_data[1]

    if input_temp > output_temp:
        stream_type = "Hot"
    else:
        stream_type = "Cold"

    #Store 1st column (input_temp) of temperature_data list in T_in list
    T_in.append(temperature_data[0])
    #Store 2nd column (output_temp) of temperature_data list in T_out list
    T_out.append(temperature_data[1])
    #Store stream_type in Type list
    Type.append(stream_type)
    #Store HX duty in Q list
    Q.append(QCALC)

print('Data Extraction Completed')

#Part 2: Heat Exchanger Pairing
print('Pairing Heat Exchangers...')
#Sort into heaters and coolers based on stream type
heaters = []
coolers = []

for i, HX_block in enumerate(HX):
    if Type[i] == 'Hot':
        heaters.append(HX_block)
    elif Type[i] == 'Cold':
        coolers.append(HX_block)

#Find HX pairs with matching duties 
pairs = []

#Create dictionaries 
heater_duties = {}
heater_inlet_temperatures = {}
heater_outlet_temperatures = {}
cooler_duties = {}
cooler_inlet_temperatures = {}
cooler_outlet_temperatures = {}

#Index heaters data
for heater in heaters:
    if heater in HX:
        i = HX.index(heater)
        heater_duties[heater] = Q[i]
        heater_inlet_temperatures[heater] = T_in[i]
        heater_outlet_temperatures[heater] = T_out[i]

#Index coolers data
for cooler in coolers:
    if cooler in HX:
        i = HX.index(cooler)
        cooler_duties[cooler] = Q[i]
        cooler_inlet_temperatures[cooler] = T_in[i]
        cooler_outlet_temperatures[cooler] = T_out[i]

#Find matching pairs
for heater in heaters:
    if heater in heater_duties:
        heater_duty = heater_duties[heater]
        for cooler in coolers:
            if cooler in cooler_duties:
                cooler_duty = cooler_duties[cooler]
                #User define tolerance value for difference in duty 
                tolerance = 1 # in %
                tolerance_fraction = tolerance/100
                #Pair HX if duty diff is within tolerance range 
                if(abs(heater_duty) > 0 and abs(cooler_duty) > 0 and abs(abs(heater_duty) - abs(cooler_duty)) / max(abs(heater_duty), abs(cooler_duty), 1e-6) < tolerance_fraction):
                    pairs.append((heater, cooler))
                    break 

Paired_heaters = []
Paired_coolers =[]
Inlet_Temperature_Hot = []
Outlet_Temperature_Hot = []
Inlet_Temperature_Cold = []
Outlet_Temperature_Cold = []
dT_Hot = []
dT_Cold = []
HX_duty = []

#Find the hot an cold end temperature differences for each heat exchanger pair 
for pair in pairs:
    heater = pair[0]
    cooler = pair[1]
    TH_in = heater_inlet_temperatures.get(heater)
    TH_out = heater_outlet_temperatures.get(heater)
    Q_heater = heater_duties.get(heater)
    TC_in = cooler_inlet_temperatures.get(cooler)
    TC_out = cooler_outlet_temperatures.get(cooler)
    Q_cooler = cooler_duties.get(cooler)
    Hot_end_dT = TH_in - TC_out
    Cold_end_dT = TH_out -TC_in
    Q_average = (abs(abs(Q_heater) + abs(Q_cooler)))/2
    Paired_heaters.append(heater)
    Paired_coolers.append(cooler)
    Inlet_Temperature_Hot.append(TH_in)
    Outlet_Temperature_Hot.append(TH_out)
    Inlet_Temperature_Cold.append(TC_in)
    Outlet_Temperature_Cold.append(TC_out)
    dT_Hot.append(Hot_end_dT)
    dT_Cold.append(Cold_end_dT)
    HX_duty.append(Q_average)

print('Pairing Completed')

#Export data to Excel 
print('Exporting Data...') 
df1 = pd.DataFrame({'HX': HX, 'T_in (°C)': T_in, 'T_out (°C)': T_out, 'Type': Type, 'Duty (kW)': Q})
df2 = pd.DataFrame({'Heater': Paired_heaters, 'Cooler': Paired_coolers, 'TH-in (°C)': Inlet_Temperature_Hot, 'TH_out (°C)': Outlet_Temperature_Hot, 'TC_in (°C)': Inlet_Temperature_Cold, 'TC_out (°C)': Outlet_Temperature_Cold, 'Hot End ΔT (°C)': dT_Hot, 'Cold End ΔT (°C)': dT_Cold, 'Duty (kW)': HX_duty})
with pd.ExcelWriter('Heat_Integration.xlsx', engine='openpyxl') as writer:
    df1.to_excel(writer, sheet_name='Extracted_Data', index=False)
    df2.to_excel(writer, sheet_name='HX_Pairs', index=False)
os.startfile('Heat_Integration.xlsx')
print('Export Completed')












           


   



  


          




           
   






    
    
    









 




    


