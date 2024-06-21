from PCAN_API.PCANBasic import *        ## PCAN-Basic library import
import tkinter as tk
import customtkinter
from tkinter import messagebox
from tkinter import IntVar

class CANFrame(customtkinter.CTkFrame):
    def __init__(self, parent, main_app):
        super().__init__(parent,fg_color="#ffffff")

        self.main_app = main_app

        self.m_objPCANBasic = PCANBasic()
        self.m_PcanHandle = PCAN_NONEBUS

        self.m_ReadingRDB = IntVar()
        self.m_ReadingRDB.set(0)

        # create sidebar frame with widgets
        self.sidebar_frame = customtkinter.CTkFrame(self, width=140, corner_radius=0)
        self.sidebar_frame.pack(side="left", fill="y")

        self.logo_label = customtkinter.CTkLabel(self.sidebar_frame, text="BMS-ADE", font=customtkinter.CTkFont(size=20, weight="bold"))
        self.logo_label.pack(padx=20, pady=(20, 10))

        welcome_message = """\
        Welcome to the ADE Battery Management System!
        Our state-of-the-art system ensures optimal 
        performance and safety for your battery operations.
        Monitor, control, and maintain your battery 
        systems efficiently and effectively.
        Navigate through CAN and RS232 interfaces with ease.
        Enhance battery performance and longevity 
        at your fingertips.
        Can frame
        """

        self.welcome_label = customtkinter.CTkLabel(self.sidebar_frame, text=welcome_message, font=customtkinter.CTkFont(size=12))
        self.welcome_label.pack(padx=20, pady=(10, 20), anchor="w")

        # Label to indicate CAN selection
        self.can_label = customtkinter.CTkLabel(self, text="CAN selected", font=customtkinter.CTkFont(size=20))
        self.can_label.pack(padx=20, pady=(40, 20))

        self.btnInit = customtkinter.CTkButton(self, width=8, text="Initialize", command= self.btnInit_Click)
        self.btnInit.pack(padx=20, pady=(20, 20))

        # Back button to navigate back to the main frame
        self.back_button = customtkinter.CTkButton(self, text="Back", command=self.go_back)
        self.back_button.pack(padx=20, pady=(20, 20))

        self.InitializeBasicComponents()

    def go_back(self):
        self.pack_forget()
        self.main_app.main_frame.pack(fill="both", expand=True)
    
    ## Button btnInit handler
    ##
    def btnInit_Click(self):
        # gets the connection values
        #
        baudrate = self.m_BAUDRATES['500 kBit/sec']
        hwtype = self.m_HWTYPES['ISA-82C200']
        ioport = int('0100',16)
        interrupt = int(3)
        print(baudrate,"baudrate")

        # Connects a selected PCAN-Basic channel
        result =  self.m_objPCANBasic.Initialize(self.m_PcanHandle,baudrate,hwtype,ioport,interrupt)

        if result != PCAN_ERROR_OK:
            if result != PCAN_ERROR_CAUTION:
                messagebox.showinfo("Error!", self.GetFormatedError(result))
            else:
                self.IncludeTextMessage('******************************************************')
                self.IncludeTextMessage('The bitrate being used is different than the given one')
                self.IncludeTextMessage('******************************************************')
                result = PCAN_ERROR_OK
        else:
            # Prepares the PCAN-Basic's PCAN-Trace file
            #
            self.ConfigureTraceFile()

        # Sets the connection status of the form
        #
        #self.SetConnectionStatus(result == PCAN_ERROR_OK)

    def GetFormatedError(self, error):
        # Gets the text using the GetErrorText API function
        # If the function success, the translated error is returned. If it fails,
        # a text describing the current error is returned.
        #
        stsReturn = self.m_objPCANBasic.GetErrorText(error, 0)
        if stsReturn[0] != PCAN_ERROR_OK:
            return "An error occurred. Error-code's text ({0:X}h) couldn't be retrieved".format(error)
        else:
            return stsReturn[1]

    def ConfigureTraceFile(self):
        # Configure the maximum size of a trace file to 5 megabytes
        #
        iBuffer = 5
        stsResult = self.m_objPCANBasic.SetValue(self.m_PcanHandle, PCAN_TRACE_SIZE, iBuffer)
        if stsResult != PCAN_ERROR_OK:
            self.IncludeTextMessage(self.GetFormatedError(stsResult))

        # Configure the way how trace files are created: 
        # * Standard name is used
        # * Existing file is ovewritten, 
        # * Only one file is created.
        # * Recording stopts when the file size reaches 5 megabytes.
        #
        iBuffer = TRACE_FILE_SINGLE | TRACE_FILE_OVERWRITE
        stsResult = self.m_objPCANBasic.SetValue(self.m_PcanHandle, PCAN_TRACE_CONFIGURE, iBuffer)        
        if stsResult != PCAN_ERROR_OK:
            self.IncludeTextMessage(self.GetFormatedError(stsResult))
        
    def IncludeTextMessage(self, strMsg):
        self.lbxInfo.insert(END, strMsg);
        self.lbxInfo.see(END)

    def SetConnectionStatus(self, bConnected=True):
        # Gets the status values for each case
        #
        self.m_Connected = bConnected
        if bConnected:
            stsConnected = "normal"
            stsNotConnected = "disabled"
        else:
            stsConnected = "disabled"
            stsNotConnected = "normal"
            
        # Buttons
        #
        self.btnInit['state'] = stsNotConnected

        # ComboBoxs
        #
        self.cbbBaudrates['state'] = stsNotConnected;
        self.cbbChannel['state'] = stsNotConnected;
        self.cbbHwType['state'] = stsNotConnected;
        self.cbbIO['state'] = stsNotConnected;
        self.cbbInterrupt['state'] = stsNotConnected;

        if ENABLE_CAN_FD:
            # Check-Buttons
            #
            self.chbCanFD['state'] = stsNotConnected;
        
        # Hardware configuration and read mode
        #
        if not bConnected:
            self.cbbChannel_SelectedIndexChanged(self.cbbChannel['value'])
            self.tmrDisplayManage(False)
        else:
            self.rdbTimer_CheckedChanged()
            self.tmrDisplayManage(True)

    def InitializeBasicComponents(self):
        self.exit = -1        
        self.m_objPCANBasic = PCANBasic()
        self.m_PcanHandle = PCAN_NONEBUS
        self.m_LastMsgsList = []

        self.m_IsFD = False
        self.m_CanRead = False
        
        self.m_NonPnPHandles = {'PCAN_ISABUS1':PCAN_ISABUS1, 'PCAN_ISABUS2':PCAN_ISABUS2, 'PCAN_ISABUS3':PCAN_ISABUS3, 'PCAN_ISABUS4':PCAN_ISABUS4, 
                                'PCAN_ISABUS5':PCAN_ISABUS5, 'PCAN_ISABUS6':PCAN_ISABUS6, 'PCAN_ISABUS7':PCAN_ISABUS7, 'PCAN_ISABUS8':PCAN_ISABUS8, 
                                'PCAN_DNGBUS1':PCAN_DNGBUS1}

        self.m_BAUDRATES = {'1 MBit/sec':PCAN_BAUD_1M, '800 kBit/sec':PCAN_BAUD_800K, '500 kBit/sec':PCAN_BAUD_500K, '250 kBit/sec':PCAN_BAUD_250K,
                            '125 kBit/sec':PCAN_BAUD_125K, '100 kBit/sec':PCAN_BAUD_100K, '95,238 kBit/sec':PCAN_BAUD_95K, '83,333 kBit/sec':PCAN_BAUD_83K,
                            '50 kBit/sec':PCAN_BAUD_50K, '47,619 kBit/sec':PCAN_BAUD_47K, '33,333 kBit/sec':PCAN_BAUD_33K, '20 kBit/sec':PCAN_BAUD_20K,
                            '10 kBit/sec':PCAN_BAUD_10K, '5 kBit/sec':PCAN_BAUD_5K}

        self.m_HWTYPES = {'ISA-82C200':PCAN_TYPE_ISA, 'ISA-SJA1000':PCAN_TYPE_ISA_SJA, 'ISA-PHYTEC':PCAN_TYPE_ISA_PHYTEC, 'DNG-82C200':PCAN_TYPE_DNG,
                         'DNG-82C200 EPP':PCAN_TYPE_DNG_EPP, 'DNG-SJA1000':PCAN_TYPE_DNG_SJA, 'DNG-SJA1000 EPP':PCAN_TYPE_DNG_SJA_EPP}

        self.m_IOPORTS = {'0100':0x100, '0120':0x120, '0140':0x140, '0200':0x200, '0220':0x220, '0240':0x240, '0260':0x260, '0278':0x278, 
                          '0280':0x280, '02A0':0x2A0, '02C0':0x2C0, '02E0':0x2E0, '02E8':0x2E8, '02F8':0x2F8, '0300':0x300, '0320':0x320,
                          '0340':0x340, '0360':0x360, '0378':0x378, '0380':0x380, '03BC':0x3BC, '03E0':0x3E0, '03E8':0x3E8, '03F8':0x3F8}

        self.m_INTERRUPTS = {'3':3, '4':4, '5':5, '7':7, '9':9, '10':10, '11':11, '12':12, '15':15}

        # if IS_WINDOWS or (not IS_WINDOWS and ENABLE_CAN_FD):
        #     self.m_PARAMETERS = {'Device ID':PCAN_DEVICE_ID, '5V Power':PCAN_5VOLTS_POWER,
        #                          'Auto-reset on BUS-OFF':PCAN_BUSOFF_AUTORESET, 'CAN Listen-Only':PCAN_LISTEN_ONLY,
        #                          'Debugs Log':PCAN_LOG_STATUS,'Receive Status':PCAN_RECEIVE_STATUS,
        #                          'CAN Controller Number':PCAN_CONTROLLER_NUMBER,'Trace File':PCAN_TRACE_STATUS,
        #                          'Channel Identification (USB)':PCAN_CHANNEL_IDENTIFYING,'Channel Capabilities':PCAN_CHANNEL_FEATURES,
        #                          'Bit rate Adaptation':PCAN_BITRATE_ADAPTING,'Get Bit rate Information':PCAN_BITRATE_INFO,
        #                          'Get Bit rate FD Information':PCAN_BITRATE_INFO_FD, 'Get CAN Nominal Speed Bit/s':PCAN_BUSSPEED_NOMINAL,
        #                          'Get CAN Data Speed Bit/s':PCAN_BUSSPEED_DATA, 'Get IP Address':PCAN_IP_ADDRESS,
        #                          'Get LAN Service Status':PCAN_LAN_SERVICE_STATUS, 'Reception of Status Frames':PCAN_ALLOW_STATUS_FRAMES,
        #                          'Reception of RTR Frames':PCAN_ALLOW_RTR_FRAMES, 'Reception of Error Frames':PCAN_ALLOW_ERROR_FRAMES,
        #                          'Interframe Transmit Delay':PCAN_INTERFRAME_DELAY, 'Reception of Echo Frames':PCAN_ALLOW_ECHO_FRAMES,
        #                          'Hard Reset Status':PCAN_HARD_RESET_STATUS, 'Communication Direction':PCAN_LAN_CHANNEL_DIRECTION}
        # else:
        #     self.m_PARAMETERS = {'Device ID':PCAN_DEVICE_ID, '5V Power':PCAN_5VOLTS_POWER,
        #                          'Auto-reset on BUS-OFF':PCAN_BUSOFF_AUTORESET, 'CAN Listen-Only':PCAN_LISTEN_ONLY,
        #                          'Debugs Log':PCAN_LOG_STATUS}
            

 






