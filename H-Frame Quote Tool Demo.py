import tkinter as tk
import numpy as np
import pandas as pd
from tkinter import filedialog, messagebox, ttk
from tkinter import *
from tkinter.ttk import *
import re
import math
import itertools
from collections import Counter
import os

# Get the directory of the current Python script
script_directory = os.path.dirname(os.path.abspath(__file__))

# Set the current working directory to the directory of the script
os.chdir(script_directory)

def createDF():
    global df
    df = pd.DataFrame(columns=['Code', 'Description', 'Quantity', 'Price', 'Discount', 'Discount Price', 'Total'])

def Refresh():
    global df
    Emptentry = pd.DataFrame({"Code": [" "],
                            "Description": [" "],
                            "Quantity": [" "],
                            "Price": [" "],
                            "Discount": [" "],
                            "Discount Price": [" "],
                            "Total": [" "]})
    df = pd.concat([df, Emptentry])
    
    clear_data()
    tv1["column"] = list(df.columns)
    tv1["show"] = "headings"
    for column in tv1["columns"]:
        tv1.heading(column, text = column)
        
    df_rows = df.to_numpy().tolist()
    for row in df_rows:
        tv1.insert("", "end", values=row)
    
    tv1.column("Code", width = 80)
    tv1.column("Description", width = 350)
    tv1.column("Quantity", width = 50, anchor=tk.CENTER)
    tv1.column("Price", width = 65, anchor=tk.CENTER)
    tv1.column("Discount", width = 40, anchor=tk.CENTER)
    tv1.column("Discount Price", width = 65, anchor=tk.CENTER)
    tv1.column("Total", width = 80, anchor=tk.CENTER)
    return None

def File_dialog():
    filename = filedialog.askopenfilename(initialdir="/", 
                                        title="Select A File", 
                                        filetypes=(("xlsx files", "*.xlsx"),("All Files", "*.*")))
    label_file["text"] = filename
    return None
    
def Load_excel_data():
    File_dialog()
    MemberList()

    file_path = label_file["text"]
    global pricedf
    try:
        excel_filename = r"{}".format(file_path)
        pricedf = pd.read_excel(excel_filename)
        pricedf = pricedf.iloc[2:, 0:3]
    except ValueError:
        tk.messagebox.showerror("Information", "The file you have chosen is invalid.")
        return None
    except FileNotFoundError:
        tk.messagebox.showerror("Information", f"No such file as {file_path}")
        return None
    label_file.config(text = "Load successful")
    
    print(pricedf.head())
    
def getCustomerList():
    global CIDList
    CIDList = Customerdf.iloc[:,0]
    global CNameList
    CNameList = Customerdf.iloc[:,1]

    global CList
    CList = []

    for i in range(0, len(CIDList)):
        CID = str(CIDList[i])
        CName = str(CNameList[i])
        
        CString = CID + " - " + CName
        
        CList.append(CString)
    
def Load_Customer_excel_data():
    File_dialog()

    file_path = label_file["text"]
    global Customerdf 
    try:
        excel_filename = r"{}".format(file_path)
        Customerdf = pd.read_excel(excel_filename)
        Customerdf = Customerdf.iloc[:, [0, 12]]
        getCustomerList()
        updateListBox(CList)
    except ValueError:
        tk.messagebox.showerror("Information", "The file you have chosen is invalid.")
        return None
    except FileNotFoundError:
        tk.messagebox.showerror("Information", f"No such file as {file_path}")
        return None
    
    label_file.config(text = "Load successful")
    print(Customerdf.head())

def Save_Excel():
    global df
    global K8df
    global quote_weight_df
    
    ConvertToK8()
    CreateWeightDF()
    
    CombinedDF = CombineDataFrames(df, K8df)
    CombinedDF = CombineDataFrames(CombinedDF, quote_weight_df)
    
    file = filedialog.asksaveasfilename(defaultextension = ".xlsx")
    CombinedDF.to_excel(str(file))
    label_file.config(text = "File saved")
    
def clear_data():
    tv1.delete(*tv1.get_children())
    pass

def MemberList():
    global df
    global Costdf
    global RafterList
    global Rafterdf
    global R138RafterList
    global R138Rafterdf
    global R162Rafterdf
    global R162RafterList
    

    Costdf = pd.read_excel("Carport Member Rates.xlsx")

    Rafterdf = Costdf.loc[:, ['Rafter Code', 'Rafter Description', 'Rafter Weight [kg/m]']]
    Rafterdf = Rafterdf.dropna()
    R162Rafterdf = Rafterdf
    RafterList = Rafterdf.iloc[:,0].tolist()
    R162RafterList = RafterList

    R138Rafterdf = Costdf.loc[:, ['R138 Rafter Code', 'R138 Rafter Description', 'R138 Rafter Weight [kg/m]']]
    R138Rafterdf = R138Rafterdf.dropna()
    R138RafterList = R138Rafterdf.iloc[:,0].tolist()
    
    global SHS100df
    global SHS100CodeList
    SHS100df = Costdf.loc[:, ['100x100 SHS Code', '100x100 SHS Description', '100x100 SHS Weight [kg/m]']]
    SHS100df = SHS100df.dropna()
    SHS100CodeList = SHS100df.iloc[:, 0].tolist()
    
    global SHS76df                                                                                                  
    global SHS76CodeList
    SHS76df = Costdf.loc[:, ['76x76 SHS Code', '76x76 SHS Description', '76x76 SHS Weight [kg/m]']]
    SHS76df = SHS76df.dropna()
    SHS76CodeList = SHS76df.iloc[:,0].tolist()
    
    global weightpm100
    weightpm100 = float(SHS100df.iloc[0,2])
    
    global weightpm76
    weightpm76 = float(SHS76df.iloc[0,2])

def round_up(number):
    """
    Rounds a number up to the nearest integer.

    Args:
        number (float): The number to round up.

    Returns:
        int: The rounded up integer.
    """
    return math.ceil(number)

def extract_percentage_value(percentage_str):
    """
    Extracts the numeric value from a percentage string and returns it as an integer.

    Args:
        percentage_str (str): A string containing a percentage (e.g., "10%").

    Returns:
        int: The numeric value without the percentage sign.
    """
    return int(percentage_str.strip().replace('%', ''))

def getSmalls():
    global SmallSmalls
    global ConSmalls
    global SuppSmalls
    
    SmallSmalls = 1 + extract_percentage_value(SSmallsVar.get())/100
    ConSmalls = 1 + extract_percentage_value(ConSmallsVar.get())/100
    SuppSmalls = 1 + extract_percentage_value(SuppSmallsVar.get())/100

def getInputs():

    createDF()
    
    #get Numerical Inputs
    global sysnum
    global pHor
    global pVer
    global pWidth
    global pLength
    global GroundClearance
    global pNum
    global angle
    global PanelO

    sysnum = int(TableNumberE.get())
    PanelO = str(OrientationVar.get())
    pHor = int(HorPanelE.get())
    pVer = int(VertPanelE.get())
    pWidth = float(PanelWidthE.get()) 
    pLength = float(PanelLengthE.get())
    angle = np.deg2rad(float(AngleE.get()))
    GroundClearance = float(GroundClearanceE.get())

    #Vert Panels
    global pVert
    pVert = int(VertPanelE.get())
    pNum = pVert*pHor

    #get ROH
    global ROH
    ROHS = var.get()
    
    if(ROHS == '600mm'):
        ROH = 600
    elif(ROHS == '800mm'):
        ROH = 800

    #get Discount 
    global discount
    discount = 1 - float(DiscountE.get())/100

    #Table Tilt
    global TableTilt
    TableTilt = MountVar.get()

    #Single or double access
    global SorD
    SorD = SDVar.get()
        
    #Rafter Splice
    global RaftSplice
    RaftSplice = RaftSVar.get()

    #Front Rafter Overhang
    global FROHang
    FROHang = int(FRaftOvE.get())
    
    #Rear Rafter Overhang
    global RROHang
    RROHang = int(RRaftOvE.get())
    
    getSmalls()

def getRaftChoice():
    global RafterLChosen
    global Rafterdf
    global RafterDescr
    global RafterCode
    
    RafterLChosenStr = str(RaftVar.get())
    
    RafterList = Rafterdf['Rafter Description'].tolist()
    
    index = 0
    
    for i in range(2,len(RafterList)):
            if (RafterList[i] == RafterLChosenStr):
                index = i
    
    RafterDescr = RafterList[index]
    
    RafterCode = Rafterdf.iloc[index, 0]
    
    #Calculating the chosen Rafter Length
    global TotalRafterL
    Rafter1L = int(get_last_set_of_numbers(RafterCode))
    
    if(RaftSplice == 'Yes'):
        Rafter2Code = RaftVar2.get()
        Rafter2L = int(get_last_set_of_numbers(Rafter2Code))
        TotalRafterL = Rafter1L + Rafter2L
    else:
        TotalRafterL = Rafter1L
    
    RafterLChosen = TotalRafterL

def RailCalc():
    global pVert
    global RailMult
    global sysnum
    global pHor
    global pWidth
    global pLength
    global pNum
    global df

    getInputs()
    #getMount()
    
    pNum = pVert*pHor
    
    if (PanelO == 'Landscape'):
        RailMult = pVert + 1
    elif (PanelO == 'Portrait'):
        RailMult = pVert * 2

def ClampCalc():

    RailCalc()

    global sysnum
    global pVert
    global pHor
    global RailMult
    global df
    global pNum


    pNum = pVert*pHor
    
    #End Clamp Calculations + Entry

    ECMult = (pHor//20 + 1) # This is for the extra gap that's required every 20 panels in a row
    
    if (PanelO == 'Landscape'):
        EndClamps = pHor * 2 * 2 * sysnum
        
        InterClamps = (RailMult - 2)*pHor*2*sysnum
    elif (PanelO == 'Portrait'):
        EndClamps = (2*RailMult)*sysnum*ECMult
        
        #InterClamp Calculations 
        InterClamps = ((RailMult*(pHor - 1)))*sysnum
        

    AddEntry("LM-EC35-RNW", round_up(EndClamps * SmallSmalls), 0)
    AddEntry("LM-IC35-GP1-RNW", round_up(InterClamps * SmallSmalls), 0)
    
def split_value(value):
    # Calculate the larger part
    larger_part = round(value / 1.6)  # 100% + 30% + 30% = 160% => value / 1.6
    # Calculate the smaller parts
    smaller_part = round(larger_part * 0.3)
    return larger_part, smaller_part, smaller_part

def getPurlinCombination(required_length):
    # List of pre-determined aluminium member lengths (in increments of 500)
    lengths = list(range(4000, 7501, 500))
    
    # Generate all possible combinations of the lengths
    for r in range(1, len(lengths) + 1):
        for combination in itertools.combinations_with_replacement(lengths, r):
            total_length = sum(combination)
            if total_length >= required_length:
                length_counts = Counter(combination)
                return combination, dict(length_counts)
    return None, None

def find5000bays(required_length):
    bay_lengths = [5000, 2500]
    overhangs2500 = [0, 900]
    overhangs5000 = [1000, 1800]
    min_overhangs = {bay_lengths[0]: overhangs5000[0], bay_lengths[1]: overhangs2500[0]}
    max_overhangs = {bay_lengths[0]: overhangs5000[1], bay_lengths[1]: overhangs2500[1]}

    def calculate_total_length(num_7500, num_5000, start_bay, end_bay):
        min_overhang_start = min_overhangs[start_bay]
        max_overhang_start = max_overhangs[start_bay]
        min_overhang_end = min_overhangs[end_bay]
        max_overhang_end = max_overhangs[end_bay]

        total_bays_length = num_7500 * 5000 + num_5000 * 2500
        for overhang_start in range(min_overhang_start, max_overhang_start + 1):
            for overhang_end in range(min_overhang_end, max_overhang_end + 1):
                total_length = total_bays_length + overhang_start + overhang_end
                if total_length == required_length:
                    return (overhang_start, overhang_end)
        return None

    min_bays_needed = float('inf')
    optimal_combination = None

    max_num_7500 = required_length // 5000 + 1
    max_num_5000 = required_length // 2500 + 1

    for num_7500 in range(max_num_7500 + 1):
        for num_5000 in range(max_num_5000 + 1):
            if num_7500 + num_5000 == 0:
                continue

            total_bays_length = num_7500 * 5000 + num_5000 * 2500

            if total_bays_length + min(min_overhangs.values()) * 2 > required_length:
                continue

            for start_bay in (5000, 2500):
                for end_bay in (5000, 2500):
                    if (num_7500 > 0 or start_bay == 2500) and (num_5000 > 0 or start_bay == 5000):
                        if (num_7500 > 0 or end_bay == 2500) and (num_5000 > 0 or end_bay == 5000):
                            overhangs = calculate_total_length(num_7500, num_5000, start_bay, end_bay)
                            if overhangs:
                                overhang_start, overhang_end = overhangs
                                num_bays = num_7500 + num_5000
                                if num_bays < min_bays_needed:
                                    min_bays_needed = num_bays
                                    optimal_combination = (num_7500, num_5000, start_bay, end_bay, overhang_start, overhang_end)

    if optimal_combination:
        num_7500, num_5000, start_bay, end_bay, overhang_start, overhang_end = optimal_combination
        print("The required length is: "+str(required_length))
        print(f"Optimal combination:")
        print(f"5000mm bays: {num_7500}")
        print(f"2500mm bays: {num_5000}")
        print(f"Start overhang: {overhang_start}mm (for {start_bay}mm bay)")
        print(f"End overhang: {overhang_end}mm (for {end_bay}mm bay)")
        print(f"Total length (including overhangs): {required_length}mm")
        return optimal_combination
    else:
        print("No valid combination found.")
        return 0,0,0,0,0,0

def find7500Bays(required_length):
    bay_lengths = [7500, 5000]
    min_overhangs = {7500: 2200, 5000: 500}
    max_overhangs = {7500: 2700, 5000: 1800}

    def calculate_total_length(num_7500, num_5000, start_bay, end_bay):
        min_overhang_start = min_overhangs[start_bay]
        max_overhang_start = max_overhangs[start_bay]
        min_overhang_end = min_overhangs[end_bay]
        max_overhang_end = max_overhangs[end_bay]

        total_bays_length = num_7500 * 7500 + num_5000 * 5000
        for overhang_start in range(min_overhang_start, max_overhang_start + 1):
            for overhang_end in range(min_overhang_end, max_overhang_end + 1):
                total_length = total_bays_length + overhang_start + overhang_end
                if total_length == required_length:
                    return (overhang_start, overhang_end)
        return None

    min_bays_needed = float('inf')
    optimal_combination = None

    max_num_7500 = required_length // 7500 + 1
    max_num_5000 = required_length // 5000 + 1

    for num_7500 in range(max_num_7500 + 1):
        for num_5000 in range(max_num_5000 + 1):
            if num_7500 + num_5000 == 0:
                continue

            total_bays_length = num_7500 * 7500 + num_5000 * 5000

            if total_bays_length + min(min_overhangs.values()) * 2 > required_length:
                continue

            for start_bay in (7500, 5000):
                for end_bay in (7500, 5000):
                    if (num_7500 > 0 or start_bay == 5000) and (num_5000 > 0 or start_bay == 7500):
                        if (num_7500 > 0 or end_bay == 5000) and (num_5000 > 0 or end_bay == 7500):
                            overhangs = calculate_total_length(num_7500, num_5000, start_bay, end_bay)
                            if overhangs:
                                overhang_start, overhang_end = overhangs
                                num_bays = num_7500 + num_5000
                                if num_bays < min_bays_needed:
                                    min_bays_needed = num_bays
                                    optimal_combination = (num_7500, num_5000, start_bay, end_bay, overhang_start, overhang_end)

    if optimal_combination:
        num_7500, num_5000, start_bay, end_bay, overhang_start, overhang_end = optimal_combination
        #print("The required length is: "+str(required_length))
        #print(f"Optimal combination:")
        #print(f"7500mm bays: {num_7500}")
        #print(f"5000mm bays: {num_5000}")
        #print(f"Start overhang: {overhang_start}mm (for {start_bay}mm bay)")
        #print(f"End overhang: {overhang_end}mm (for {end_bay}mm bay)")
        #print(f"Total length (including overhangs): {required_length}mm")
        return optimal_combination
    else:
        print("No valid combination found.")
        return 0,0,0,0,0,0

def getPurlins():

    ClampCalc()

    global sysnum
    global pHor
    global pWidth
    global pLength
    global RafterLChosen
    global SupportLegs
    global pVert
    global df
    global RailMult
    global PurlinLMin
    
    if (PanelO == 'Landscape'):
        CalcPurlinL = pHor*pLength + 20*(pHor - 1)
    elif(PanelO == 'Portrait'):
        CalcPurlinL = pHor*pWidth + 20*(pHor - 1) + 200
    
    CalcPurlLabel.config(text = "Required Purlin Length (mm): " +str(CalcPurlinL))
    
    SuppPurlinList, SuppPurlinDict = getPurlinCombination(CalcPurlinL)
    
    #PurlinLMin = sum(SuppPurlinList)
    PurlinLMin = round_up(CalcPurlinL + 300)
    
    PurlinCode = "LM-CP-P-R160-2.5-"
    PurlinSplicesNum = 0
    
    LPQuantity = 0
    
    for Length, Lquantity in SuppPurlinDict.items():
        PurlinCode = "LM-CP-P-R160-2.5-"+str(Length)
        LPQuantity = (LPQuantity + Lquantity)
        AddEntry(PurlinCode, round_up(Lquantity*RailMult*sysnum * SuppSmalls), 0)
        
    
    PurlinSplicesNum = round_up((LPQuantity - 1)*RailMult*sysnum * SuppSmalls)
    AddEntry("LM-CP-P-RS-660", PurlinSplicesNum, 0)
    
        
    StitchingScrews = round_up(18*(PurlinSplicesNum) * SmallSmalls)
    AddEntry("FS-S-22X6-C4", StitchingScrews, 0)
   
    PurlinSuppString = "Supplied Purlin Length: " + str(PurlinLMin) + "mm"
    PurlinLabel.config(text = PurlinSuppString)
    
    # Support calculation for carports
    global bays5mcalc
    global bays7p5mcalc
    global POHang1
    global POHang2
    global POHangwarning
    
    POHang1 = 0
    POHang2 = 0
    
    POHangwarning = 0 #0 means there is no warning, 1 means there is a warning
    
    #calculating the base number of 7.5m and 5m bays that can fit in the required purlin length 
    global MaxBay
    MaxBay = MaxBVar.get()
    
    if(MaxBay == "7.5m"):
        bays7p5mcalc, bays5mcalc, start_bay, end_bay, POHang1, POHang2 = find7500Bays(PurlinLMin)
    elif(MaxBay == "5m"):
        bays7p5mcalc, bays5mcalc, start_bay, end_bay, POHang1, POHang2 = find5000bays(PurlinLMin)
    
    global message
    if(bays7p5mcalc == 0 and bays5mcalc == 0):
        message = "Please add or remove panels"
    else:
        message = "Success"
        
    global PurlinOHang
    PurlinOHang = POHang1 + POHang2

    #Calculating the number of supports needed based on the number of parking bays supplied
    global SupportLegsC
    SupportLegsC = 1 + (bays5mcalc + bays7p5mcalc)
    

    global SupportLegs
    SupportLegs = SupportLegsC * sysnum

    #Adding Purlin to Rafter Connectors
    global PRC
    totalrails = RailMult
    PRC = ((totalrails * SupportLegsC)*4)*sysnum
    
    PRCCode = "LM-PRC"
    AddEntry(PRCCode, round_up(PRC * SmallSmalls), 0)
    
    #Display the required Rafter Length
    if (PanelO == 'Landscape'):
        if(SorD == 'East-West Butterfly'):
            CalcRafterL = (pVert/2) * pWidth + (((pVert/2) - 1)*20) + 200 + 500
        else:
            CalcRafterL = pVert * pWidth + ((pVert - 1)*20) + 200
    elif (PanelO == 'Portrait'):
        if(SorD == 'East-West Butterfly'):
            CalcRafterL = (pVert/2) * pLength + (((pVert/2) - 1)*20) - ROH + 500
        else:    
            CalcRafterL = pVert * pLength + ((pVert - 1)*20) - ROH
        
    if (RaftSplice == 'Yes'):
        secondRafterEntry()
    else:
        RafterChoiceLabel.config(text = "Please select a standard Rafter Lengths in mm:")
    
    selection = "Calculated Rafter length: "+str(CalcRafterL)
    CalcRaftLabel.config(text = selection)
    
    SupportString = "7.5m bays: " + str(bays7p5mcalc) + "\n5m bays: " + str(bays5mcalc)
    SupportSLabel.config(text = SupportString)

    PurlinSuppString = "Supplied Purlin Length: " + str(PurlinLMin) + "mm"
    PurlinLabel.config(text = PurlinSuppString)
    
    SupportLegsStr = "Number of Support Legs per structure: " + str(SupportLegsC)
    SupportLegsLabel.config(text = SupportLegsStr)
    
    OHangString = "Purlin Overhang 1: "+str(POHang1)+"mm. \nPurlin Overhang 2: " + str(POHang2) + "mm"
    OHangLabel.config(text = OHangString)
    
    selection = "Calculated Rafter length: "+str(CalcRafterL)
    CalcRaftLabel.config(text = selection)

def get_last_set_of_numbers(input_string):
    # Find all sequences of digits in the input string
    matches = re.findall(r'\d+', input_string)
    
    # Return the last match if there are any, otherwise return None
    return matches[-1] if matches else None

def secondRafterEntry():
    RafterChoiceLabel.config(text = "Please select your rafters: ")
    
    global RaftVar2
    RaftVar2 = tk.StringVar()
    RaftStr2 = Rafterdf['Rafter Code'].tolist()
    RaftVar2.set(RaftStr2[0])
    RafterChoice2Op = tk.OptionMenu(InputFrame, RaftVar2, *RaftStr2)
    RafterChoice2Op.grid(row = 11, column = 3, padx = 5, pady = 5)

def RafterEntry(quantity):
    global df
    global RafterCode
        
    if (RaftSplice == 'Yes'):
        global Rafter2Code
        
        Rafter2Code = RaftVar2.get()
        RafterCode = RaftVar.get()
        
        r1description, r1price, r1discprice, r1total = getprice(RafterCode, quantity, 0)
        discountp = float(DiscountE.get())
        RaftEntry = pd.DataFrame({"Code": [RafterCode], 
                                "Description": [r1description],
                                "Quantity": [quantity], 
                                "Price": [r1price],
                                "Discount": [str(discountp) + "%"],
                                "Discount Price": [r1discprice],
                                "Total": [r1total]})
        df = pd.concat([df, RaftEntry])
        
        r2description, r2price, r2discprice, r2total = getprice(Rafter2Code, quantity, 0)
        discountp = float(DiscountE.get())
        RaftEntry = pd.DataFrame({"Code": [Rafter2Code], 
                                "Description": [r2description],
                                "Quantity": [quantity], 
                                "Price": [r2price],
                                "Discount": [str(discountp) + "%"],
                                "Discount Price": [r2discprice],
                                "Total": [r2total]})
        df = pd.concat([df, RaftEntry])
        
        raftersplicenum = quantity
        AddEntry("LM-CP-SB-MILL-L", quantity, 1000)
        FSHBM16x150 = quantity * 6
        AddEntry("FS-HB-M16X150", FSHBM16x150, 0)
        FSFWM16 = FSHBM16x150 * 2
        AddEntry("FS-FW-M16", FSFWM16, 0)
        FSSWM16 = FSHBM16x150
        AddEntry("FS-SW-M16", FSSWM16, 0)
        FSNM16 = FSHBM16x150
        AddEntry("FS-N-M16", FSNM16, 0)
        
    elif (RaftSplice == 'No'):
        RafterCode = RaftVar.get()
        
        r1description, r1price, r1discprice, r1total = getprice(RafterCode, quantity, 0)
        discountp = float(DiscountE.get())
        RaftEntry = pd.DataFrame({"Code": [RafterCode], 
                                "Description": [r1description],
                                "Quantity": [quantity], 
                                "Price": [r1price],
                                "Discount": [str(discountp) + "%"],
                                "Discount Price": [r1discprice],
                                "Total": [r1total]})
        df = pd.concat([df, RaftEntry])
        
        if (SorD == 'East-West Butterfly'):
            #Adding East-West Connector plates and bolts
            EWCPlates = quantity
            AddEntry("LM-CP-EWC", EWCPlates, 0)
            FSHBM16x150 = quantity * 2
            AddEntry("FS-HB-M16X150", FSHBM16x150, 0)
            FSFWM16 = FSHBM16x150 * 2
            AddEntry("FS-FW-M16", FSFWM16, 0)
            FSSWM16 = FSHBM16x150
            AddEntry("FS-SW-M16", FSSWM16, 0)
            FSNM16 = FSHBM16x150
            AddEntry("FS-N-M16", FSNM16, 0)
            
    #Calculating the chosen Rafter Length
    global TotalRafterL
    Rafter1L = int(get_last_set_of_numbers(RafterCode))
    
    if(RaftSplice == 'Yes'):
        Rafter2L = int(get_last_set_of_numbers(Rafter2Code))
        TotalRafterL = Rafter1L + Rafter2L
    else:
        TotalRafterL = Rafter1L

def MountSupp():

    global sysnum
    global SupportLegs
    global pVert
    global RafterLChosen
    global RailMult
    global df
    global RafterLChosen
    global Rafterdf
    global RafterDescr
    global RafterCode
    
    if (SorD == 'East-West Butterfly'):
        RafterQuantity = SupportLegs*2
    else:
        RafterQuantity = SupportLegs
    
    RafterEntry(round_up(RafterQuantity * SuppSmalls))

    global FPOhang
    FPOhang = 500
        
    if(PanelO == 'Landscape'):
        FPOhang = 0
    
    #Calculates the length of the support bars by taking the panel and rafter overhang into account
    global FrontSupport
    global RearSupport
    global FPSpacing
    global PlinthHeight
    global RemainingRafterL
    
    FPSpacingS = FPSpacingVar.get()
    
    if (FPSpacingS == '1.0 m single-access'):
        Hframestr = "1"
    elif (FPSpacingS == '1.2 m single-access'):
        Hframestr = "1.2"
    elif (FPSpacingS == '1.5 m double-access'):
        Hframestr = "1.5"
    elif (FPSpacingS == '2.0 m double-access'):
        Hframestr = "2.0"
    FPSpacing = float(FPSpacingS[0:3]) * 1000
    PlinthHeight = 1850
    RemainingRafterL = TotalRafterL - FROHang - RROHang
    CPIB = 0
        
    if (SorD == "Single-access"): #Determining if it's a single or double access structure
        
        #Adding H-Frame Base
        CPIB = SupportLegs
        
        if (TableTilt == "Standard Tilt"):
            AddEntry("LMK-CP-HB-S-S-"+Hframestr, round_up(CPIB * SuppSmalls), 0)
            M12x35 = round_up(SupportLegs * 4 * SmallSmalls)
            AddEntry("FS-HB-M12X35", M12x35, 0)
            FWM12 = M12x35 * 2
            AddEntry("FS-FW-M12", FWM12, 0)
            SWM12 = M12x35
            AddEntry("FS-SW-M12", SWM12, 0)
            NM12 = M12x35
            AddEntry("FS-N-M12", NM12, 0)
            
            if(TotalRafterL == 6200 and angle == np.deg2rad(10) and GroundClearance == 2500 and FROHang == 1200 and RROHang == 350 and FPSpacing == 1200):
                RearSupport = 1075
                RearSupportCode = "LM-CP-SB-MILL-"+str(RearSupport)
                CentreSupport = 1180
                CentreSupportCode = "LM-CP-SB-MILL-"+str(CentreSupport)
                FrontSupport = 3670
                FrontSupportCode = "LM-CP-SB-MILL-"+str(FrontSupport)
                
                AddEntry(RearSupportCode, round_up(SupportLegs * SuppSmalls), 0)
                AddEntry(CentreSupportCode, round_up(SupportLegs * SuppSmalls), 0)
                AddEntry(FrontSupportCode, round_up(SupportLegs * SuppSmalls), 0)
            
            elif(TotalRafterL == 6500 and angle == np.deg2rad(10) and GroundClearance == 2500 and FROHang == 1200 and RROHang == 350 and FPSpacing == 1200):
                RearSupport = 1075
                RearSupportCode = "LM-CP-SB-MILL-"+str(RearSupport)
                CentreSupport = 1180
                CentreSupportCode = "LM-CP-SB-MILL-"+str(CentreSupport)
                FrontSupport = 3940
                FrontSupportCode = "LM-CP-SB-MILL-"+str(FrontSupport)
                
                AddEntry(RearSupportCode, round_up(SupportLegs * SuppSmalls), 0)
                AddEntry(CentreSupportCode, round_up(SupportLegs * SuppSmalls), 0)
                AddEntry(FrontSupportCode, round_up(SupportLegs * SuppSmalls), 0)
            
            else:
                RearL = GroundClearance + RROHang*np.sin(angle)
                u1 = RearL - PlinthHeight
                RearSupport = round(((u1)**2 + FPSpacing**2)**0.5)
                
                u2 = u1 + (RemainingRafterL*0.42)*np.sin(angle)
                d2 = (RemainingRafterL*0.42)*np.cos(angle) - FPSpacing
                CentreSupport = round((u2**2 + d2**2)**0.5)
                
                u3 = u1 + RemainingRafterL*np.sin(angle)
                d3 = RemainingRafterL*np.cos(angle) - FPSpacing
                FrontSupport = round((u3**2 + d3**2)**0.5)
                
                #adding the support bars
                AddEntry("LM-CP-SB-MILL-L", round_up(SupportLegs * SuppSmalls), FrontSupport)
                AddEntry("LM-CP-SB-MILL-L", round_up(SupportLegs * SuppSmalls), CentreSupport)
                AddEntry("LM-CP-SB-MILL-L", round_up(SupportLegs * SuppSmalls), RearSupport)


        elif (TableTilt == "Reverse Tilt"):
            AddEntry("LMK-CP-HB-S-R-"+Hframestr, round_up(CPIB * SuppSmalls), 0)
            #adding M12x35 HB for the H-frame base assembly
            M12x35 = round_up(SupportLegs * 4 * SmallSmalls)
            AddEntry("FS-HB-M12X35", M12x35, 0)
            FWM12 = M12x35 * 2
            AddEntry("FS-FW-M12", FWM12, 0)
            SWM12 = M12x35
            AddEntry("FS-SW-M12", SWM12, 0)
            NM12 = M12x35
            AddEntry("FS-N-M12", NM12, 0)
            
            if(TotalRafterL == 6200 and angle == np.deg2rad(10) and GroundClearance == 2500 and FROHang == 1200 and RROHang == 350 and FPSpacing == 1200):
                RearSupport = 1740
                RearSupportCode = "LM-CP-SB-MILL-"+str(RearSupport)
                CentreSupport = 1450
                CentreSupportCode = "LM-CP-SB-MILL-"+str(CentreSupport)
                FrontSupport = 3400
                FrontSupportCode = "LM-CP-SB-MILL-"+str(FrontSupport)
                
                AddEntry(RearSupportCode, round_up(SupportLegs * SuppSmalls), 0)
                AddEntry(CentreSupportCode, round_up(SupportLegs * SuppSmalls), 0)
                AddEntry(FrontSupportCode, round_up(SupportLegs * SuppSmalls), 0)
            
            elif(TotalRafterL == 6500 and angle == np.deg2rad(10) and GroundClearance == 2500 and FROHang == 1200 and RROHang == 350 and FPSpacing == 1200):
                RearSupport = 1780
                RearSupportCode = "LM-CP-SB-MILL-"+str(RearSupport)
                CentreSupport = 1450
                CentreSupportCode = "LM-CP-SB-MILL-"+str(CentreSupport)
                FrontSupport = 3715
                FrontSupportCode = "LM-CP-SB-MILL-"+str(FrontSupport)
                
                AddEntry(RearSupportCode, round_up(SupportLegs * SuppSmalls), 0)
                AddEntry(CentreSupportCode, round_up(SupportLegs * SuppSmalls), 0)
                AddEntry(FrontSupportCode, round_up(SupportLegs * SuppSmalls), 0)
                
            else:
                RearL = GroundClearance + (TotalRafterL - RROHang)*np.sin(angle)
                u1 = RearL - PlinthHeight
                d1 = FPSpacing
                RearSupport = round((d1**2 + u1**2)**0.5)
                
                u3 = GroundClearance + FROHang*np.sin(angle)
                d3 = RemainingRafterL*np.cos(angle) - FPSpacing
                FrontSupport = round((u3**2 + d3**2)**0.5)
                
                CSRL = RemainingRafterL*0.42
                d2 = CSRL*np.cos(angle) - FPSpacing
                u2 = u3 + (RemainingRafterL - CSRL)*np.sin(angle)
                CentreSupport = round((d2**2 + u2**2)**0.5)
                
                #adding the support bars
                AddEntry("LM-CP-SB-MILL-L", round_up(SupportLegs * SuppSmalls), FrontSupport)
                AddEntry("LM-CP-SB-MILL-L", round_up(SupportLegs * SuppSmalls), CentreSupport)
                AddEntry("LM-CP-SB-MILL-L", round_up(SupportLegs * SuppSmalls), RearSupport)

        #Adding the connectors for the support bars
        RLC1 = round_up(SupportLegs*2 * ConSmalls)
        AddEntry("LM-CP-RLC-1", RLC1, 0)
        RLC2 = round_up(SupportLegs * ConSmalls)
        AddEntry("LM-CP-RLC-2", RLC2, 0)
        FP3 = round_up(SupportLegs * ConSmalls)
        AddEntry("LM-CP-FP-3", FP3, 0)
        FFW = round_up(FP3*2 * ConSmalls)
        AddEntry("LM-CP-FFW", FFW, 0)
        
        #Adding the M24x65 bolts to anchor the foot pieces to the H-frame base 
        M24x65 = round_up((FP3) * 2 * SmallSmalls)
        AddEntry("FS-HB-M24X65", M24x65, 0)
        FWM24 = M24x65 * 2
        AddEntry("FS-FW-M24", FWM24, 0)
        NM24 = M24x65
        AddEntry("FS-N-M24", NM24, 0)
        
    elif(SorD == "Double-access"):
        #Adding H-Frame Base
        CPIB = SupportLegs
        AddEntry("LMK-CP-HB-D", round_up(CPIB * SuppSmalls), 0)
        
        if(TotalRafterL == 10800 and angle == np.deg2rad(8) and GroundClearance == 2500 and FROHang == 1200 and RROHang == 1200 and FPSpacing == 1500):
            RearSupport = 3400
            RearSupportCode = "LM-CP-SB-MILL-"+str(RearSupport)
            CentreSupport1 = 1300
            CentreSupportCode1 = "LM-CP-SB-MILL-"+str(CentreSupport1)
            CentreSupport2 = 1370
            CentreSupportCode2 = "LM-CP-SB-MILL-"+str(CentreSupport2)
            CentreSupport3 = 1370
            CentreSupportCode3 = "LM-CP-SB-MILL-"+str(CentreSupport3)
            CentreSupport4 = 1690
            CentreSupportCode4 = "LM-CP-SB-MILL-"+str(CentreSupport4)
            FrontSupport = 3820
            FrontSupportCode = "LM-CP-SB-MILL-"+str(FrontSupport)
                
            AddEntry(RearSupportCode, round_up(SupportLegs * SuppSmalls), 0)
            AddEntry(CentreSupportCode1, round_up(SupportLegs * SuppSmalls), 0)
            AddEntry(CentreSupportCode2, round_up(SupportLegs * SuppSmalls), 0)
            AddEntry(CentreSupportCode3, round_up(SupportLegs * SuppSmalls), 0)
            AddEntry(CentreSupportCode4, round_up(SupportLegs * SuppSmalls), 0)
            AddEntry(FrontSupportCode, round_up(SupportLegs * SuppSmalls), 0)
                
        elif(TotalRafterL == 11200 and angle == np.deg2rad(8) and GroundClearance == 2500 and FROHang == 1200 and RROHang == 1200 and FPSpacing == 1500):
            RearSupport = 3497
            RearSupportCode = "LM-CP-SB-MILL-"+str(RearSupport)
            CentreSupport1 = 1300
            CentreSupportCode1 = "LM-CP-SB-MILL-"+str(CentreSupport1)
            CentreSupport2 = 1370
            CentreSupportCode2 = "LM-CP-SB-MILL-"+str(CentreSupport2)
            CentreSupport3 = 1370
            CentreSupportCode3 = "LM-CP-SB-MILL-"+str(CentreSupport3)
            CentreSupport4 = 1690
            CentreSupportCode4 = "LM-CP-SB-MILL-"+str(CentreSupport4)
            FrontSupport = 4020
            FrontSupportCode = "LM-CP-SB-MILL-"+str(FrontSupport)
                
            AddEntry(RearSupportCode, round_up(SupportLegs * SuppSmalls), 0)
            AddEntry(CentreSupportCode1, round_up(SupportLegs * SuppSmalls), 0)
            AddEntry(CentreSupportCode2, round_up(SupportLegs * SuppSmalls), 0)
            AddEntry(CentreSupportCode3, round_up(SupportLegs * SuppSmalls), 0)
            AddEntry(CentreSupportCode4, round_up(SupportLegs * SuppSmalls), 0)
            AddEntry(FrontSupportCode, round_up(SupportLegs * SuppSmalls), 0)
            
        else:
        
            u1 = GroundClearance + RROHang*np.sin(angle) - PlinthHeight
            d1 = (TotalRafterL/2)*np.cos(angle) - (FPSpacing/2)
            RearSupport = round(((u1)**2 + (d1)**2)**0.5)
            
            CSRL = 0.35 * (RemainingRafterL/2)
            d2 = CSRL*np.cos(angle) - (FPSpacing/2)
            u2 = u1 + ((RemainingRafterL/2 - CSRL)*np.sin(angle))
            CentreSupport1 = round((u2**2 + d2**2)**0.5)
            
            u3 = u2 + CSRL*np.sin(angle)
            d3 = FPSpacing/2
            CentreSupport2 = round((u3**2 + d3**2)**0.5)
            
            CentreSupport3 = CentreSupport2
            
            u4 = u3 + CSRL*np.sin(angle)
            d5 = CSRL*np.cos(angle) - (FPSpacing/2)
            CentreSupport4 = round((u4**2 + d5**2)**0.5)
            
            u5 = u3 + (RemainingRafterL/2)*np.sin(angle)
            d6 = (RemainingRafterL/2)*np.cos(angle) - (FPSpacing/2)
            FrontSupport = round((u5**2 + d6**2)**0.5)
            
            #adding the support bars
            AddEntry("LM-CP-SB-MILL-L", round_up(SupportLegs * SuppSmalls), FrontSupport)
            AddEntry("LM-CP-SB-MILL-L", round_up(SupportLegs * SuppSmalls), CentreSupport1)
            AddEntry("LM-CP-SB-MILL-L", round_up(SupportLegs * SuppSmalls), CentreSupport2)
            AddEntry("LM-CP-SB-MILL-L", round_up(SupportLegs * SuppSmalls), CentreSupport3)
            AddEntry("LM-CP-SB-MILL-L", round_up(SupportLegs * SuppSmalls), CentreSupport4)
            AddEntry("LM-CP-SB-MILL-L", round_up(SupportLegs * SuppSmalls), RearSupport)

        #Adding the connectors for the support bars
        RLC1 = round_up(SupportLegs*4 * ConSmalls)
        AddEntry("LM-CP-RLC-1", RLC1, 0)
        RLC2 = round_up(SupportLegs * ConSmalls)
        AddEntry("LM-CP-RLC-2", RLC2, 0)
        FP3 = round_up(SupportLegs*2 * ConSmalls)
        AddEntry("LM-CP-FP-3", FP3, 0)
        FFW = round_up(FP3*2 * ConSmalls)
        AddEntry("LM-CP-FFW", FFW, 0)
        
        #Adding the M24x65 bolts to anchor the foot pieces to the H-frame base 
        M24x65 = round_up((FP3) * 2 * SmallSmalls)
        AddEntry("FS-HB-M24X65", M24x65, 0)
        FWM24 = M24x65 * 2
        AddEntry("FS-FW-M24", FWM24, 0)
        NM24 = M24x65
        AddEntry("FS-N-M24", NM24, 0)
        
    elif (SorD == 'East-West Butterfly'):
        #Adding H-Frame Base
        CPIB = SupportLegs
        AddEntry("LMK-CP-HB-D-EW", round_up(CPIB * SuppSmalls), 0)

        if(TotalRafterL == 6500 and angle == np.deg2rad(6) and GroundClearance == 2500 and FROHang == 1200 and FPSpacing == 2000):
            RearSupport = 4440
            RearSupportCode = "LM-CP-SB-MILL-"+str(RearSupport)
            CentreSupport1 = 1450
            CentreSupportCode1 = "LM-CP-SB-MILL-"+str(CentreSupport1)
            CentreSupport2 = 1100
            CentreSupportCode2 = "LM-CP-SB-MILL-"+str(CentreSupport2)
            CentreSupport3 = 1100
            CentreSupportCode3 = "LM-CP-SB-MILL-"+str(CentreSupport3)
            CentreSupport4 = 1450
            CentreSupportCode4 = "LM-CP-SB-MILL-"+str(CentreSupport4)
            FrontSupport = 4440
            FrontSupportCode = "LM-CP-SB-MILL-"+str(FrontSupport)
                
            AddEntry(RearSupportCode, round_up(SupportLegs * SuppSmalls), 0)
            AddEntry(CentreSupportCode1, round_up(SupportLegs * SuppSmalls), 0)
            AddEntry(CentreSupportCode2, round_up(SupportLegs * SuppSmalls), 0)
            AddEntry(CentreSupportCode3, round_up(SupportLegs * SuppSmalls), 0)
            AddEntry(CentreSupportCode4, round_up(SupportLegs * SuppSmalls), 0)
            AddEntry(FrontSupportCode, round_up(SupportLegs * SuppSmalls), 0)
            
        else:
            u3 = GroundClearance - PlinthHeight
            d3 = FPSpacing/2
            CSRL = (TotalRafterL - FROHang)*0.35
            CentreSupport2 = round((u3**2 + d3**2)**0.5)
            
            CentreSupport3 = CentreSupport2
            
            u2 = u3 + CSRL*np.sin(angle)
            d2 = CSRL*np.cos(angle) - (FPSpacing/2)
            CentreSupport1 = round((u2**2 + d2**2)**0.5)
            
            CentreSupport4 = CentreSupport1
            
            d1 = (TotalRafterL - FROHang)*np.cos(angle) - d3
            u1 = u3 + (TotalRafterL - FROHang)*np.sin(angle)
            FrontSupport = round((u1**2 + d1**2)**0.5)
            
            RearSupport = FrontSupport
            
            #adding the support bars
            AddEntry("LM-CP-SB-MILL-L", round_up(SupportLegs * SuppSmalls), FrontSupport)
            AddEntry("LM-CP-SB-MILL-L", round_up(SupportLegs * SuppSmalls), CentreSupport1)
            AddEntry("LM-CP-SB-MILL-L", round_up(SupportLegs * SuppSmalls), CentreSupport2)
            AddEntry("LM-CP-SB-MILL-L", round_up(SupportLegs * SuppSmalls), CentreSupport3)
            AddEntry("LM-CP-SB-MILL-L", round_up(SupportLegs * SuppSmalls), CentreSupport4)
            AddEntry("LM-CP-SB-MILL-L", round_up(SupportLegs * SuppSmalls), RearSupport)

        #Adding the connectors for the support bars
        RLC1 = round_up(SupportLegs*6 * ConSmalls)
        AddEntry("LM-CP-RLC-1", RLC1, 0)
        FP3 = round_up(SupportLegs*2 * ConSmalls)
        AddEntry("LM-CP-FP-3", FP3, 0)
        FFW = round_up(FP3*2 * ConSmalls)
        AddEntry("LM-CP-FFW", FFW, 0)
        
        #Adding the M24x65 bolts to anchor the foot pieces to the H-frame base 
        M24x65 = round_up((FP3) * 2 * SmallSmalls)
        AddEntry("FS-HB-M24X65", M24x65, 0)
        FWM24 = M24x65 * 2
        AddEntry("FS-FW-M24", FWM24, 0)
        NM24 = M24x65
        AddEntry("FS-N-M24", NM24, 0)
        
            
    #Adding the Threaded rods for Bases
    FSTRM16x200 = round_up((CPIB) * 8 * SmallSmalls)
    AddEntry("FS-TR-M16x200-HDG", FSTRM16x200, 0)
    
    FSFWM16 = FSTRM16x200 * 2
    AddEntry("FS-FW-M16-HDG", FSFWM16, 0)
    
    FSNM16 = FSTRM16x200 * 2
    AddEntry("FS-N-M16-HDG", FSNM16, 0)
    
    #Calculating Purlin End Caps
    PurECs = round_up(RailMult * 2 * sysnum * SmallSmalls)
    AddEntry("LM-CP-PEC", PurECs, 0)
        
    #Calculating Rafter End Caps
    RaftECMult = 1
    if (SorD == 'East-West Butterfly'):
        RaftECMult = 2
    RaftECs = round_up(2 * SupportLegs * RaftECMult * SmallSmalls)
    RECCode = "LM-CP-REC"
    AddEntry(RECCode, RaftECs, 0)
        
    #Adding Stitching Screws for End Caps
    ECStitchScr = round_up((PurECs + RaftECs) * 2 * SmallSmalls)
    AddEntry("FS-S-22X6-C4", ECStitchScr, 0)
                
    #Getting the cross-bracing connections:
    CB5m = int(CB5mE.get())
    if(CB5m is None):
        CB5m = 0
    CB7p5m = int(CB7p5mE.get())
    if(CB7p5m is None):
        CB7p5m = 0
    
    BracingType = TCBVar.get()
    
    DAMult = 1
    
    FP90 = 0
    FP120 = 0
    FSTRM12x160 = 0
    
    
    if (BracingType == "X-Bracing"):
        CBSupportBars6m = (2 * CB7p5m + 2 * CB5m) * sysnum
        AddEntry("LM-SB-6000", round_up(CBSupportBars6m * SuppSmalls), 0)
        
        TTC90 = round_up((CB7p5m * 2 + CB5m * 2) * sysnum * ConSmalls)
        AddEntry("LMK-TTC90-CBC", TTC90, 0)
        FSHBM16x150 = round_up(TTC90  * SmallSmalls)
        AddEntry("FS-HB-M16X150", FSHBM16x150, 0)
        FSFWM16 = FSHBM16x150 * 2
        AddEntry("FS-FW-M16", FSFWM16, 0)
        FSSWM16 = FSHBM16x150
        AddEntry("FS-SW-M16", FSSWM16, 0)
        FSNM16 = FSHBM16x150
        AddEntry("FS-N-M16", FSNM16, 0)
        
        CBPlate = round_up((2*CB7p5m + 2*CB5m) * sysnum * ConSmalls)
        AddEntry("LM-CP-CBC", CBPlate, 0)
        FSHBM10x90 = round_up((CB7p5m * 4 + CB5m * 4) * sysnum * SmallSmalls)
        AddEntry("FS-HB-M10X90", FSHBM10x90, 0)
        FSFWM10 = FSHBM10x90 * 2
        AddEntry("FS-FW-M10", FSFWM10, 0)
        FSSWM10 = FSHBM10x90
        AddEntry("FS-SW-M10", FSSWM10, 0)
        FSNM10 = FSHBM10x90
        AddEntry("FS-N-M10", FSNM10, 0)
        
        FP90 = round_up((CB7p5m * 2 + CB5m * 2) * sysnum * ConSmalls)
        AddEntry("LM-FP-90", FP90, 0)
        FSTRM12x160 = round_up(FP90 * 2 * SmallSmalls)
        AddEntry("FS-TR-M12X160", FSTRM12x160, 0)
        FSFWM12 = FSTRM12x160 * 2
        AddEntry("FS-FW-M12", FSFWM12, 0)
        FSNM12 = FSTRM12x160 * 2
        AddEntry("FS-N-M12", FSNM12, 0)
        
    
    elif (BracingType == "V-Bracing"):
        if(CB5m > 0):
            CBSupportBars6m = (2 * CB5m) * sysnum
            AddEntry("LM-SB-6000", round_up(CBSupportBars6m * SuppSmalls), 0)
        if(CB7p5m > 0):
            CBSupportBars5m = (2 * CB7p5m) * sysnum
            AddEntry("LM-SB-5000", round_up(CBSupportBars5m * SuppSmalls), 0)
        
        TTC90 = round_up((CB7p5m + CB5m) * 2 * sysnum * ConSmalls)
        AddEntry("LMK-TTC90-CBC", TTC90, 0)
        FSHBM16x150 = round_up(TTC90 * SmallSmalls)
        AddEntry("FS-HB-M16X150", FSHBM16x150, 0)
        FSFWM16 = FSHBM16x150 * 2
        AddEntry("FS-FW-M16", FSFWM16, 0)
        FSSWM16 = FSHBM16x150
        AddEntry("FS-SW-M16", FSSWM16, 0)
        FSNM16 = FSHBM16x150
        AddEntry("FS-N-M16", FSNM16, 0)
        
        if(CB5m > 0):
            CBPlate = round_up(2 * CB5m * sysnum * ConSmalls)
            AddEntry("LM-CP-CBC", CBPlate, 0)    
            FSHBM10x90 = round_up(4 * CB5m * sysnum * SmallSmalls)
            AddEntry("FS-HB-M10X90", FSHBM10x90, 0)
            FSFWM10 = 2 * FSHBM10x90
            AddEntry("FS-FW-M10", FSFWM10, 0)
            FSSWM10 = FSHBM10x90
            AddEntry("FS-SW-M10", FSSWM10, 0)
            FSNM10 = FSHBM10x90
            AddEntry("FS-N-M10", FSNM10, 0)
            
            FP90 = round_up(CB5m * 2 * sysnum * ConSmalls)
            AddEntry("LM-FP-90", FP90, 0)
            
        if(CB7p5m > 0):
            FP120 = round_up(CB7p5m * 1 * sysnum * ConSmalls)
            AddEntry("LM-FP-120", FP120, 0)
            
        FSTRM12x160 = round_up((FP90 + FP120) * 2 * SmallSmalls)
        AddEntry("FS-TR-M12X160", FSTRM12x160, 0)
        FSFWM12 = FSTRM12x160 * 2
        AddEntry("FS-FW-M12", FSFWM12, 0)
        FSNM12 = FSTRM12x160 * 2
        AddEntry("FS-N-M12", FSNM12, 0)
        
    elif (BracingType == "Rafter Bracing"):
        if(SorD == "Double-access"):
            DAMult = 2
            
        if(CB5m > 0):
            CBSupportBars6m = (4 * CB5m) * DAMult * sysnum
            AddEntry("LM-SB-L", round_up(CBSupportBars6m * SuppSmalls), 3600)
        if(CB7p5m > 0):
            CBSupportBars5m = (4 * CB7p5m) * DAMult * sysnum
            AddEntry("LM-SB-L", round_up(CBSupportBars5m * SuppSmalls), 4600)
            
        TTC90 = round_up((CB7p5m * 4 + CB5m * 4) * DAMult * sysnum * ConSmalls)
        AddEntry("LMK-TTC90-CBC", TTC90, 0)
        FSHBM16x150 = round_up(TTC90 * SmallSmalls)
        AddEntry("FS-HB-M16X150", FSHBM16x150, 0)
        FSFWM16 = FSHBM16x150 * 2
        AddEntry("FS-FW-M16", FSFWM16, 0)
        FSSWM16 = FSHBM16x150
        AddEntry("FS-SW-M16", FSSWM16, 0)
        FSNM16 = FSHBM16x150
        AddEntry("FS-N-M16", FSNM16, 0)
        
        CBPlate = round_up((2 * CB7p5m + 2 * CB5m) * DAMult * sysnum * ConSmalls)
        AddEntry("LM-CP-CBC", CBPlate, 0)
        FSHBM10x90 = round_up((4 * CB7p5m + 4 * CB5m) * DAMult * sysnum * SmallSmalls)
        AddEntry("FS-HB-M10X90", FSHBM10x90, 0)
        FSFWM10 = FSHBM10x90 * 2
        AddEntry("FS-FW-M10", FSFWM10, 0)
        FSSWM10 = FSHBM10x90
        AddEntry("FS-SW-M10", FSSWM10, 0)
        FSNM10 = FSHBM10x90
        AddEntry("FS-N-M10", FSNM10, 0)
        
    IKA70007 = ((FSTRM12x160)//20 + FSTRM16x200//20 + 2)
    AddEntry("IKA-70007", IKA70007, 0)

def replace_first_l_with_numbers(input_str, replacement_numbers):
    count = 0
    result = ''

    for char in input_str:
        if char == 'L':
            count += 1
            if count == 1:
                result += str(replacement_numbers)  # Replace 'L' with the desired numbers
            else:
                result += char
        else:
            result += char

    return result

def getprice(code, quantity, length):
    global pricedf
    global discountp
    global RafterLChosen
    
    discountp = float(DiscountE.get())

    discount = 1 - discountp/100

    ref = pricedf.iloc[:,0]
    prices = pricedf.iloc[:,2]
    descriptions = pricedf.iloc[:, 1]
    
    string = code

    index = 0

    if (code == "LM-R110-4200"):
        price = 1214.36 
        RafterLChosen = 4200
        description = "Rafter 110x4200mm AL6005 T6 Mill"
    else:
        for i in range(2,len(ref)):
            if (ref[i] == string):
                index = i
                
        price = prices[index]
        description = descriptions[index]
    
    if(code == "LM-CP-SB-MILL-L"):
        pricet = (price*(length/1000)+40)
        description = "Carport Support Bar 118x" + str(length) + "mm AL6063 T6 Mill"
    elif(code == "LM-SB-L"):
        pricet = (float(price)*(length/1000)+17)
        description = "Support Bar 55x55x" + str(length) + "mm AL6063 T6 Mill"
    else:
        pricet = price

    price = round(float(pricet), 2)
    discprice = round(pricet*discount, 2)   
    totalprice = discprice*quantity
     
    
    return description, price, discprice, totalprice

def extract_economax_length(text):
    match = re.search(r'LENGTH:\s*(\d+)', text)
    if match:
        return int(match.group(1))
    return None

def extract_length(s: str) -> int:
    
    if("Carport Support Bar" in s):
        match = re.search(r'\d+x(\d+)mm', s)

    else:
        match = re.search(r'\d+x\d+x(\d+)mm', s)
        
    if match:
        return int(match.group(1))
    raise ValueError("Length not found in string")    
    
def getStdSupportBarLength(Description):
    
    length = extract_length(Description)
    
    if("Carport Support Bar" in Description):
        SBLengthArray = [1075, 1100, 1145, 1180, 1300, 1370, 1450, 1500, 1690, 1710, 
                         1740, 1780, 2130, 2280, 2320, 2440, 2480, 2510, 2530, 2580, 
                         2750, 2980, 3400, 3440, 3497, 3580, 3670, 3715, 3820, 3940, 
                         3990, 4020, 4050, 4180, 4270, 4360, 4440, 4580, 4760, 6000]
    else:
        SBLengthArray = [450, 530, 550, 590, 615, 1500, 1550, 1560, 1740, 1800, 
                         1815, 1870, 1875, 1960, 2525, 2540, 2610, 2615, 2670, 2710, 
                         2795, 2820, 2840, 2890, 3000, 3340, 5000, 6000]
    
    if (length <= SBLengthArray[0]):
        return SBLengthArray[0]
    
    else:
        for i in range(1, len(SBLengthArray)):
            if length <= SBLengthArray[i] and length > SBLengthArray[i-1]:
                return SBLengthArray[i]

def AddK8Entry(code, quantity):
    global K8df
    
    NewEntry = pd.DataFrame({"Code": [code],
                             "Quantity": [str(quantity)]})
    K8df = pd.concat([K8df, NewEntry])

def ConvertToK8():
    global K8df
    
    K8df = pd.DataFrame(columns=['Code', 'Quantity'])
    
    K8Convertdf = pd.read_excel("Old and New Codes.xlsx")

    Oldcodedf = K8Convertdf.loc[:, ['Old Code']]
    Oldcodedf = Oldcodedf.dropna()
    OldcodeList = Oldcodedf.iloc[:,0].tolist()
    
    Newcodedf = K8Convertdf.loc[:, ['New Code']]
    Newcodedf = Newcodedf.dropna()
    NewcodeList = Newcodedf.iloc[:,0].tolist()
    
    QuoteCodesdf = df.loc[:, ['Code']]
    QuoteCodesdf = QuoteCodesdf.dropna()
    QuoteCodes = QuoteCodesdf.iloc[:,0].tolist()
    
    QuoteDescdf = df.loc[:, ['Description']]
    QuoteDescdf = QuoteDescdf.dropna()
    QuoteDescs = QuoteDescdf.iloc[:,0].tolist() 
    
    QuoteQuantitiesdf = df.loc[:, ['Quantity']]
    QuoteQuantitiesdf = QuoteQuantitiesdf.dropna()
    QuoteQuantities = QuoteQuantitiesdf.iloc[:,0].tolist()
    
    for i in range(0, len(QuoteCodes) - 1):
        for j in range(0, len(OldcodeList) - 1):
            
            if (QuoteCodes[i] == "LM-SB-L"):
                length = getStdSupportBarLength(QuoteDescs[i])
                StdSuppBarCode = "LM-SB-" + str(length)
                QuoteCodes[i] = StdSuppBarCode
                
            elif (QuoteCodes[i] == "LM-CP-SB-MILL-L"):
                length = getStdSupportBarLength(QuoteDescs[i])
                if (length == 6000):
                    StdSuppBarCode = "LM-CP-SB-6000"
                else:
                    StdSuppBarCode = "LM-CP-SB-MILL-" + str(length)
                QuoteCodes[i] = StdSuppBarCode 
                     
            if (QuoteCodes[i] == OldcodeList[j]):
                
                AddK8Entry(NewcodeList[j], QuoteQuantities[i])

def LoadWeights():
    global Weightdf
    
    Weightdf = pd.read_excel("Inventory Volume & weight.xlsx")
    Weightdf = Weightdf.iloc[8:, 0:3].reset_index(drop=True)
    Weightdf.columns = Weightdf.iloc[0]
    Weightdf = Weightdf.iloc[1:, :].reset_index(drop=True)
    Weightdf = Weightdf.iloc[:, [0, 2]]
    Weightdf = Weightdf.dropna()
    
    global WeightCode
    global Weights
    WeightCode = Weightdf.iloc[:,0].tolist()
    Weights = Weightdf.iloc[:,1].tolist()

def getWeight(code, description, quantity):
    global WeightCode
    global Weights
    
    for i in range(0, len(WeightCode) - 1):
        if ("LM-CP-EC-" in code):
            Length = extract_economax_length(description) #not dividing by 1000 because the weightpm is kg/m and the weights are provided in grams
                
            if Length is None:
                print("Length not found in description: ", description)
                
            if("100" in code):
                weight = round(weightpm100 * Length)
            elif("76" in code):
                weight = round(weightpm76 * Length)
            else:
                print("Unknown code for Economax Support: ", code)
                weight = 0
                
        elif (code == WeightCode[i]):
            if (code == "LM-SB-L"):
                Length = extract_length(description)/1000
                weight = round(Weights[i] * Length)
            elif (code == "LM-CP-SB-MILL-L"):
                Length = extract_length(description)/1000
                weight = round(Weights[i] * Length)
            else:
                weight = Weights[i]
            break
        else:
            weight = 0
    
    TotWeight = weight * int(quantity)
    TotWeight = round(float(TotWeight))
            
    return weight, TotWeight

def AddWeightEntry(weight, TotWeight):
    global quote_weight_df
    
    NewEntry = pd.DataFrame({"Unit Weight [g]": [float(weight)],
                             "Total Weight [g]": [float(TotWeight)]})
    quote_weight_df = pd.concat([quote_weight_df, NewEntry])

def CreateWeightDF():
    LoadWeights()
    
    global df
    global Weightdf
    
    # Create a new DataFrame with the same index as df
    global quote_weight_df
    quote_weight_df = pd.DataFrame(columns=['Unit Weight [g]', 'Total Weight [g]'])
    
    for i in range(1, len(df) - 1):
        code = df.iloc[i]['Code']
        description = df.iloc[i]['Description']
        quantity = df.iloc[i]['Quantity']
        
        weight, TotWeight = getWeight(code, description, quantity)
        
        # Add the weights to the new DataFrame
        AddWeightEntry(weight, TotWeight)
    
    #Adding total weight of the order    
    OrderWeight = (quote_weight_df['Total Weight [g]'].sum())/1000
    OrderWeight = round(float(OrderWeight), 3)
    OrderWeight = str(OrderWeight) + " kg"
    #AddWeightEntry("Total weight of the order", OrderWeight)
    NewEntry = pd.DataFrame({"Unit Weight [g]": ["Total weight of the order"],
                             "Total Weight [g]": [str(OrderWeight)]})
    quote_weight_df = pd.concat([NewEntry, quote_weight_df])

def CombineDataFrames(df1: pd.DataFrame, df2: pd.DataFrame) -> pd.DataFrame:
    global df
    global K8df
    
    # Combine the two DataFrames
    df1 = df1.reset_index(drop=True)
    df2 = df2.reset_index(drop=True)
    
    return pd.concat([df1, df2], axis=1, ignore_index=False)

def AddEntry(code, quantity, length):
    global df
    global discountp
    
    description, price, discprice, total = getprice(code, quantity, length)
    
    NewEntry = pd.DataFrame({"Code": [code], 
                            "Description": [str(description)],
                            "Quantity": [quantity], 
                            "Price": [price],
                            "Discount": [str(discountp)+"%"],
                            "Discount Price": [discprice],
                            "Total": [total]})
    df = pd.concat([df, NewEntry])

def Calculations():
    debug = "Lol"
    debugLabel.config(text = debug)
    getPurlins()

def getDescription():
    global df
    global sysnum
    global angle
    global GroundClearance
    global pHor
    global pVert
    global pLength
    global pWidth
    global PurlinLMin
    global SupportLegsC
    
    description = "Table details: Table count: "+str(sysnum)+" Alu-MAX H-Frame, "+str(SorD)+", "+str(pVer)+"x"+str(pHor)+", "+str(PanelO)+", Support Count: "+str(SupportLegsC)
    description = description + ", Table length: "+str(PurlinLMin)+"mm, Purlin Runs: "+str(RailMult)+", for panel dimensions: "+str(pLength)+"x"+str(pWidth)+"mm, with "+str(bays5mcalc)+" 5m bays"
    description = description +" and "+str(bays7p5mcalc)+" 7.5m bays."
    
    total = round((df['Total'].sum()), 2)
    
    Descrentry = pd.DataFrame({"Code": ["DESCRIPTION"], 
                            "Description": [description],
                            "Quantity": [sysnum], 
                            "Price": [" "],
                            "Discount": [str(discountp)+"%"],
                            "Discount Price": [" "],
                            "Total": [total]})
    df = pd.concat([Descrentry, df])
    
def FinishCalc():
    
    getRaftChoice()
    MountSupp()
    getDescription()

    debugLabel.config(text = message)
    
def updateListBox(data):
    # clear list box
    ClientListBox.delete(0, END)
    
    # Add Clients to list box
    for item in data:
        ClientListBox.insert(END, item)
        
#Update entry box with listbox clicked
def fillout(e):
    #delete whatever is in the entry box
    CCodeE.delete(0, END)
    
    # Add clicked list item to entry box
    CCodeE.insert(0, ClientListBox.get(ACTIVE))
    
# Create function to check entry vs listbox
def check(e):
    # grab what was typed
    typed = CCodeE.get()
    
    if typed =='':
        data = CList
        updateListBox(data)
    else:
        data = []
        for item in CList:
            if typed.lower() in item.lower():
                data.append(item)
    
    updateListBox(data)

def ProjectInfo():
    # Toplevel object which will 
    # be treated as a new window
    global newWindow
    newWindow = Toplevel(root)
 
    # sets the title of the
    # Toplevel widget
    newWindow.title("Project Information Entry")
 
    # sets the geometry of toplevel
    newWindow.geometry("1380x750")
    
    #Customer List collect
    CustomerListFrame = tk.LabelFrame(newWindow, text = "Load the latest customer list")
    CustomerListFrame.grid(row = 0, column = 0, padx = 5, pady = 5)
    
    #Customer Information
    CCodeLabel = tk.Label(CustomerListFrame, text = "Customer Details:")
    CCodeLabel.grid(row = 1, column = 0, padx = 5, pady = 5)
    global CCodeE
    CCodeE = tk.Entry(CustomerListFrame, width = 75)
    CCodeE.grid(row = 1, column = 1, padx = 5, pady = 5)
    global ClientListBox
    ClientListBox = tk.Listbox(CustomerListFrame, width = 75)
    ClientListBox.grid(row = 2, column = 1, padx = 5, pady = 5)
    
    # Create a binding on the listbox on click
    ClientListBox.bind("<<ListboxSelect>>", fillout)
    
    # Create a binding on the entry box
    CCodeE.bind("<KeyRelease>", check)
    
    #Button to find customer file
    CustomerListB = tk.Button(CustomerListFrame, text = "Load Customer List", command = lambda: Load_Customer_excel_data())
    CustomerListB.grid(row = 0, column = 0, padx = 5, pady = 5)
    
    #Project Details Entries
    PDFrame  = tk.LabelFrame(newWindow, text = "Enter the project details")
    PDFrame.grid(row = 0, column = 1, padx = 5, pady = 5)
    
    #Date
    DateLabel = tk.Label(PDFrame, text = "Please enter today's date (YYYY/MM/DD):")
    DateLabel.grid(row = 0, column = 0, padx = 5, pady = 5)
    global DateE
    DateE = tk.Entry(PDFrame)
    DateE.grid(row = 0, column = 1, padx = 5, pady = 5)
    
    #Reference
    ReferenceLabel = tk.Label(PDFrame, text = "Enter Quote Reference:")
    ReferenceLabel.grid(row = 1, column = 0, padx = 5, pady = 5)
    global ReferenceE
    ReferenceE = tk.Entry(PDFrame, width = 75)
    ReferenceE.grid(row = 1, column = 1, padx = 5, pady = 5)
    
    #Message
    MessageLabel = tk.Label(PDFrame, text = "Enter Quote Message:")
    MessageLabel.grid(row = 2, column = 0)
    global MessageE
    MessageE = tk.Entry(PDFrame, width = 75)
    MessageE.grid(row = 2, column = 1, padx = 5, pady = 5)

    #Buttons
    ButtonFrame  = tk.LabelFrame(newWindow, text = "Capture Information")
    ButtonFrame.grid(row = 3, column = 0, padx = 5, pady = 5)
    
    PIButton = tk.Button(ButtonFrame, text = "Capture Project Information", command = lambda: getProjectInfo())
    PIButton.grid(row = 0, column = 0, padx = 5, pady = 5)
    
def getProjectInfo():
    
    global transaction
    transaction = 'Quote'
    
    global date
    date = str(DateE.get())
    
    global QuoteRef
    QuoteRef = ReferenceE.get()
    
    global QuoteMessage
    QuoteMessage  = MessageE.get()
    
    global CustomerCode
    CCode = CCodeE.get()
    Customerarray = CCode.split('-')
    CustomerCode = Customerarray[0]
    
    global termname
    termname = 'CASH'
    
    global state
    state = 'Pending'
    
    global WarehouseID
    WarehouseID = 'Lumax - Olifantsfontein'
    
    global Unit
    Unit = 'Each'
    
    global DepartmentID
    DepartmentID = 'GroundMounting'
    
    global Sodcust
    Sodcust = 'CASH'
    
    newWindow.destroy()

def CreateSageImport():
    global df
    
    # Step 1: Read the quote template CSV file into a DataFrame
    template_file = 'import template.csv'
    template_df = pd.read_csv(template_file)

    # Display the template DataFrame
    #print("Template DataFrame:")
    #print(template_df.head())

    # Step 2: Assume you have a quote DataFrame with new data
    # Example quote DataFrame (replace this with your actual quote DataFrame)
    #quote_file = 'test.xlsx'
    #quote_df = pd.read_excel(quote_file)
    quote_df = df
    quote_df = quote_df.loc[:, ['Code', 'Description', 'Quantity', 'Discount Price']]
    quote_df.rename(columns={'Code': 'ITEMID', 'Description': 'ITEMDESC', 'Quantity': 'QUANTITY', 'Discount Price': 'PRICE'}, inplace=True)
                                            
    # Display the quote DataFrame
    #print("\nQuote DataFrame:")
    #print(quote_df.head())

    # Step 3: Create a list of dictionaries to represent rows for the merged DataFrame
    merged_data = []

    # Add the first row of the template DataFrame as the header row in the merged DataFrame
    merged_data.append(dict(zip(template_df.columns, template_df.iloc[0])))

    # Iterate over each row in the quote DataFrame and map it to the template columns
    for _, quote_row in quote_df.iterrows():
        # Create a dictionary to hold data for the new row
        new_row = {}

        # Map the quote data to the corresponding template columns
        for col in quote_df.columns:
            if col in template_df.columns:
                new_row[col] = quote_row[col]  # Assign quote data to the corresponding template column

        # Append the new row dictionary to the list
        merged_data.append(new_row)

    # Create the merged DataFrame directly from the list of dictionaries
    global merged_df
    merged_df = pd.DataFrame(merged_data, columns=template_df.columns)

    #Updating the line 1 items:
    merged_df.at[1, 'TRANSACTIONTYPE'] = transaction

    merged_df.at[1, 'DATE'] = date

    merged_df.at[1, 'GLPOSTINGDATE'] = date

    merged_df.at[1, 'CUSTOMER_ID'] = CustomerCode

    merged_df.at[1, 'TERMNAME'] = termname

    merged_df.at[1, 'REFERENCENO'] = QuoteRef

    merged_df.at[1, 'MESSAGE'] = QuoteMessage

    merged_df.at[1, 'STATE'] = state

    for i in range(1, (len(merged_df.index) - 1)):
        merged_df.at[i, 'LINE'] = i
        merged_df.at[i, 'WAREHOUSEID'] = WarehouseID
        merged_df.at[i, 'UNIT'] = "Each"
        merged_df.at[i, 'DEPARTMENTID'] = DepartmentID
        merged_df.at[i, 'LOCATIONID'] = "100 - Lumax"
        merged_df.at[i, 'SODOCUMENTENTRY_CUSTOMERID'] = Sodcust
        
    # Display the merged DataFrame
    #print("\nMerged DataFrame:")
    #print(merged_df)

    # Step 4: Save the merged DataFrame to a new CSV file
    Save_CSV()

def Save_CSV():
    global merged_df
    file = filedialog.asksaveasfilename(defaultextension = ".csv")
    merged_df.to_csv(str(file), index=False)
    label_file.config(text = "File saved")
    
root = tk.Tk()
root.geometry("1380x750")
root.title("H-Frame Quote Tool (Use at your own risk)")

InputFrame = tk.LabelFrame(root, text = "Table Data Entry: ")
InputFrame.pack(side = "top", fill = "x")
#InputFrame.place(height = 400, width = 1380)

DispFrame = tk.LabelFrame(root, text = "Calculated Quote: ")
DispFrame.pack(expand = True,fill = "both")
#DispFrame.place(height = 350, width = 1380, rely = 0.525, relx = 0)

LoadPricesB = tk.Button(InputFrame, text = "Load current Prices", command = lambda: Load_excel_data())
LoadPricesB.grid(row = 1, column = 1, padx = 5, pady = 5)

CustomerInfoB = tk.Button(InputFrame, text = "Enter Project Info", command = lambda: ProjectInfo())
CustomerInfoB.grid(row = 1, column = 3, padx = 5, pady = 5)

label_file = ttk.Label(InputFrame, text = "No file selected")
label_file.grid(row = 1, column = 2, padx = 5, pady = 5)

debugLabel = tk.Label(InputFrame, text = "Lol")
debugLabel.grid(row = 1, column = 5, padx = 5, pady = 5)

SupportLabel = tk.Label(InputFrame, text = "This is a simple demo.")
SupportLabel.grid(row = 1, column = 6, padx = 5, pady = 5)

TableNumberLabel = tk.Label(master = InputFrame, text = "Number of Tables:")
TableNumberLabel.grid(row = 2, column = 1, padx = 5, pady = 5)
TableNumberE = tk.Entry(InputFrame)
TableNumberE.grid(row = 2, column = 2, padx = 5, pady = 5)

MountLabel = tk.Label(InputFrame, text = "Table Tilt:")
MountLabel.grid(row = 2, column = 3, padx = 5, pady = 5)
MountVar = tk.StringVar()
MountStr = ['Standard Tilt', 'Reverse Tilt']
MountVar.set(MountStr[0])
MountOp = tk.OptionMenu(InputFrame, MountVar, *MountStr)
MountOp.grid(row = 2, column = 4, padx = 5, pady = 5)

VertPanelLabel = tk.Label(InputFrame, text = "Table Width (no. of panels):")
VertPanelLabel.grid(row = 4, column = 1, padx = 5, pady = 5)
VertPanelE = tk.Entry(InputFrame)
VertPanelE.grid(row = 4, column = 2, padx = 5, pady = 5)

OrientationLabel = tk.Label(InputFrame, text = "Panel Orientation:")
OrientationLabel.grid(row = 3, column = 1, padx = 5, pady = 5)
OrientationVar = tk.StringVar()
OrientationList = ['Portrait', 'Landscape']
OrientationVar.set(OrientationList[0])
OrientationOp = tk.OptionMenu(InputFrame, OrientationVar, *OrientationList)
OrientationOp.grid(row = 3, column = 2, padx=5, pady=5)

HorPanelLabel = tk.Label(InputFrame, text = "Table Length (no. of panels):")
HorPanelLabel.grid(row = 5, column = 1, padx = 5, pady = 5)
HorPanelE = tk.Entry(InputFrame)
HorPanelE.grid(row = 5, column = 2, padx = 5, pady = 5)

ROHLabel = tk.Label(InputFrame, text = "Please select a total panel overhang on rafter:")
ROHLabel.grid(row = 6, column = 1, padx = 5, pady = 5)
var = tk.StringVar()
RaftOvList = ['600mm', '800mm']
var.set(RaftOvList[0])
RaftOvOp = tk.OptionMenu(InputFrame, var, *RaftOvList)
RaftOvOp.grid(row = 6, column = 2, padx = 5, pady = 5)

DiscountLabel = tk.Label(InputFrame, text = "Customer Discount [%]:")
DiscountLabel.grid(row = 7, column = 1, padx = 5, pady = 5)
DiscountE = tk.Entry(InputFrame)
DiscountE.grid(row = 7, column = 2, padx = 5, pady = 5)

PanelWidthLabel = tk.Label(InputFrame, text = "Width of the selected panels:")
PanelWidthLabel.grid(row = 3, column = 3, padx = 5, pady = 5)
PanelWidthE = tk.Entry(InputFrame)
PanelWidthE.grid(row = 3, column = 4, padx = 5, pady = 5)

PanelLengthLabel = tk.Label(InputFrame, text = "Length of the selected panels:")
PanelLengthLabel.grid(row = 4, column = 3, padx = 5, pady = 5)
PanelLengthE = tk.Entry(InputFrame)
PanelLengthE.grid(row = 4, column = 4, padx = 5, pady = 5)

AngleLabel = tk.Label(InputFrame, text = "Angle (degrees):")
AngleLabel.grid(row = 5, column = 3, padx = 5, pady = 5)
AngleE = tk.Entry(InputFrame)
AngleE.grid(row = 5, column = 4, padx = 5, pady = 5)

GroundClearanceLabel = tk.Label(InputFrame, text = "Ground Clearance:")
GroundClearanceLabel.grid(row = 6, column = 3, padx = 5, pady = 5)
GroundClearanceE = tk.Entry(InputFrame)
GroundClearanceE.grid(row = 6, column = 4, padx = 5, pady = 5)

SDLabel = tk.Label(InputFrame, text = "Single or Double-access:")
SDLabel.grid(row = 7, column = 3, padx = 5, pady = 5)
SDVar = tk.StringVar()
SDList = ['Single-access', 'Double-access', 'East-West Butterfly']
SDVar.set(SDList[0])
SDOp = tk.OptionMenu(InputFrame, SDVar, *SDList)
SDOp.grid(row = 7, column = 4, padx = 5, pady = 5)

FPSpacingLabel = tk.Label(InputFrame, text = "Foot Piece Spacing:")
FPSpacingLabel.grid(row = 8, column = 3, padx = 5, pady = 5)
FPSpacingVar = tk.StringVar()
FPSpacingList = ['1.0 m single-access', '1.2 m single-access', '1.5 m double-access', '2.0 m East-West']
FPSpacingVar.set(FPSpacingList[0])
FPSpacingOP = tk.OptionMenu(InputFrame, FPSpacingVar, *FPSpacingList)
FPSpacingOP.grid(row = 8, column = 4, padx = 5, pady = 5)

MemberList()
RaftSLabel = tk.Label(InputFrame, text = "Rafter Splice?")
RaftSLabel.grid(row = 2, column = 5, padx = 5, pady = 5)
RaftSVar = tk.StringVar()
RaftSList = ['No', 'Yes']
RaftSVar.set(RaftSList[0])
RaftSOp = tk.OptionMenu(InputFrame, RaftSVar, *RaftSList)
RaftSOp.grid(row = 2, column = 6, padx=5, pady=5)

FRaftOvLabel = tk.Label(InputFrame, text = "Front Rafter Overhang [mm]:") 
FRaftOvLabel.grid(row = 3, column = 5, padx = 5, pady = 5)
FRaftOvE = tk.Entry(InputFrame)
FRaftOvE.grid(row = 3, column = 6, padx = 5, pady = 5)

RRaftOvLabel = tk.Label(InputFrame, text = "Rear Rafter Overhang [mm]:")
RRaftOvLabel.grid(row = 4, column = 5, padx = 5, pady =5)
RRaftOvE = tk.Entry(InputFrame)
RRaftOvE.grid(row = 4, column = 6, padx = 5, pady = 5)

MBLabel = tk.Label(InputFrame, text = "Maximum Support Spacing:")
MBLabel.grid(row = 5, column = 5, padx = 5, pady = 5)
MaxBVar = tk.StringVar()
MBList = ['7.5m', '5m']
MaxBVar.set(MBList[0])
MBOp = tk.OptionMenu(InputFrame, MaxBVar, *MBList)
MBOp.grid(row = 5, column = 6, padx = 5, pady = 5)

CB5mLabel = tk.Label(InputFrame, text = "Number of 5m bay cross-bracing")
CB5mLabel.grid(row = 6, column = 5, padx = 5, pady = 5)
CB5mE = tk.Entry(InputFrame)
CB5mE.grid(row = 6, column = 6, padx = 5, pady = 5)

CB7p5mLabel = tk.Label(InputFrame, text = "Number of 7.5m bay cross-bracing:")
CB7p5mLabel.grid(row = 7, column = 5, padx = 5, pady = 5)
CB7p5mE = tk.Entry(InputFrame)
CB7p5mE.grid(row = 7, column = 6, padx = 5, pady = 5)

TCrossBraceLabel = tk.Label(InputFrame, text = "Type of 7.5m bay Cross-bracing\n(The Rafter bracing applies to 5m bays as well):")
TCrossBraceLabel.grid(row = 8, column = 5, padx = 5, pady = 7)
global TCBVar
TCBVar = tk.StringVar()
TCBString = ['V-Bracing', 'X-Bracing', 'Rafter Bracing']
TCBVar.set(TCBString[2])
TCBOp = tk.OptionMenu(InputFrame, TCBVar, *TCBString)
TCBOp.grid(row = 8, column = 6, padx = 5, pady = 5)

SSmallsLabel = tk.Label(InputFrame, text = "Extra Fasteners and Clamps Percentage:")
SSmallsLabel.grid(row = 9, column = 1, padx = 5, pady = 5)
global SSmallsVar
SSmallsVar = tk.StringVar()
SSmallsList = ['2%', '5%', '10%']
SSmallsVar.set(SSmallsList[0])
SSmallsOp = tk.OptionMenu(InputFrame, SSmallsVar, *SSmallsList)
SSmallsOp.grid(row = 9, column = 2, padx = 5, pady = 5)

ConSmallsLabel = tk.Label(InputFrame, text = "Extra Connectors(TTC's, FP's) Percentage:")
ConSmallsLabel.grid(row = 9, column = 3, padx = 5, pady = 5)
global ConSmallsVar
ConSmallsVar = tk.StringVar()
ConSmallsList = ['0%', '1%', '2%', '3%', '4%', '5%']
ConSmallsVar.set(ConSmallsList[0])
ConSmallsOp = tk.OptionMenu(InputFrame, ConSmallsVar, *ConSmallsList)
ConSmallsOp.grid(row = 9, column = 4, padx = 5, pady = 5)

SuppSmallsLabel = tk.Label(InputFrame, text = "Extra Supports Percentage:")
SuppSmallsLabel.grid(row = 9, column = 5, padx = 5, pady = 5)
global SuppSmallsVar
SuppSmallsVar = tk.StringVar()
SuppSmallsList = ['0%', '1%', '2%', '3%', '4%', '5%']
SuppSmallsVar.set(SuppSmallsList[0])
SuppSmallsOp = tk.OptionMenu(InputFrame, SuppSmallsVar, *SuppSmallsList)
SuppSmallsOp.grid(row = 9, column = 6, padx = 5, pady = 5)

CalcRaftB = tk.Button(InputFrame, text = "Calculate Rafter Length", command = lambda: Calculations())
CalcRaftB.grid(row = 10, column = 1, padx = 5, pady = 5)
CalcRaftLabel = tk.Label(InputFrame, text = " ")
CalcRaftLabel.grid(row = 10, column = 2, padx = 5, pady = 5)

RafterChoiceLabel = tk.Label(InputFrame, text = "Please select a Rafter Length:")
RafterChoiceLabel.grid(row = 11, column = 1, padx = 5, pady = 5)
global RaftVar
RaftVar = tk.StringVar()
#RaftStr = ['3400', '3600', '3800', '4000', '4200', '4400', '5400', '5600', '6200']
RaftStr = Rafterdf['Rafter Code'].tolist()
RaftVar.set(RaftStr[0])
RafterChoiceOp = tk.OptionMenu(InputFrame, RaftVar, *RaftStr)
RafterChoiceOp.grid(row = 11, column = 2, padx = 5, pady = 5)

CalcPurlLabel = tk.Label(InputFrame, text = "Calculated Purlin Length")
CalcPurlLabel.grid(row = 12, column = 1, padx = 5, pady = 5)

PurlinLabel = tk.Label(InputFrame, text = "Supplied Purlin Length:")
PurlinLabel.grid(row = 12, column = 2, padx = 5, pady = 5)

SupportSLabel = tk.Label(InputFrame, text = "Support Spacing")
SupportSLabel.grid(row = 12, column = 3, padx = 5, pady = 5)

SupportLegsLabel = tk.Label(InputFrame, text = "Support Legs")
SupportLegsLabel.grid(row = 12, column = 4, padx = 5, pady = 5)

OHangLabel = tk.Label(InputFrame, text = "Overhang")
OHangLabel.grid(row = 12, column = 5, padx = 5, pady = 5)

TotalPriceLabel = tk.Label(InputFrame, text = "Total Price of the quote:")
TotalPriceLabel.grid(row = 12, column = 6, padx = 5, pady = 5)

CalcQuoteB = tk.Button(InputFrame, text = "Calculate Quote", command = lambda: FinishCalc())
CalcQuoteB.grid(row = 13, column = 1, padx = 5, pady = 5)

DispQuoteB = tk.Button(InputFrame, text = "Display Quote", command = lambda: Refresh())
DispQuoteB.grid(row = 13, column = 2, padx = 5, pady = 5)

ExportB = tk.Button(InputFrame, text = "Export Quote", command = lambda: Save_Excel())
ExportB.grid(row = 13, column = 3, padx = 5, pady = 5)

SageB = tk.Button(InputFrame, text = "Create Sage Import", command = lambda: CreateSageImport())
SageB.grid(row = 13, column = 4, padx = 5, pady = 5)

# Treeview Widget
tv1 = ttk.Treeview(DispFrame)
tv1.place(relheight=1, relwidth=1)

treescrolly = tk.Scrollbar(DispFrame, orient = "vertical", command=tv1.yview)
treescrollx = tk.Scrollbar(DispFrame, orient = "horizontal", command = tv1.xview)
tv1.configure(xscrollcommand = treescrollx.set, yscrollcommand = treescrolly.set)
treescrollx.pack(side = "bottom", fill = "x")
treescrolly.pack(side = "right", fill = "y")

# Add weights to the grid rows and columns
# Changing the weights will change the size of the rows/columns relative to each other
DispFrame.grid_rowconfigure(0, weight=1)
DispFrame.grid_rowconfigure(1, weight=1)
DispFrame.grid_columnconfigure(0, weight=1)
DispFrame.grid_columnconfigure(1, weight=1)

root.mainloop()