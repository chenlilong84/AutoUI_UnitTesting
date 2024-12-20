using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PP5AutoUITests
{
    public enum TestCmdGroupType
    {
        // 第一組
        [System.ComponentModel.Description("AC Source")]
        AC_Source,

        [System.ComponentModel.Description("DC Source")]
        DC_Source,

        [System.ComponentModel.Description("Power Analyzer")]
        Power_Analyzer,

        [System.ComponentModel.Description("Digital Multimeter")]
        Digital_Multimeter,

        [System.ComponentModel.Description("XYZ Table")]
        XYZ_Table,

        [System.ComponentModel.Description("Digital Oscilloscope")]
        Digital_Oscilloscope,

        [System.ComponentModel.Description("RF Voltmeter")]
        RF_Voltmeter,

        [System.ComponentModel.Description("Electronic Load")]
        Electronic_Load,

        [System.ComponentModel.Description("TIMS")]
        TIMS,

        [System.ComponentModel.Description("I/O CARD")]
        IO_Card,

        [System.ComponentModel.Description("RS-232")]
        RS232,

        [System.ComponentModel.Description("Input Source")]
        Input_Source,

        [System.ComponentModel.Description("Timing/Noise Analyzer")]
        Timing_Noise_Analyzer,

        [System.ComponentModel.Description("I2C Device")]
        I2C_Device,

        [System.ComponentModel.Description("EEPROM")]
        EEPROM,

        [System.ComponentModel.Description("GPIB")]
        GPIB,

        [System.ComponentModel.Description("AC Load")]
        AC_Load,

        [System.ComponentModel.Description("CAN Card")]
        CAN_Card,

        [System.ComponentModel.Description("DMM / SU")]
        DMM_SU,

        [System.ComponentModel.Description("Short Circuit/OVP Tester")]
        Short_Circuit_OVP_Tester,

        [System.ComponentModel.Description("On/Off Controller")]
        On_Off_Controller,


        // 第二組
        [System.ComponentModel.Description("HSRLOAD")]
        HSRLOAD,

        [System.ComponentModel.Description("Glitch Detector")]
        Glitch_Detector,

        [System.ComponentModel.Description("Function Generator")]
        Function_Generator,

        [System.ComponentModel.Description("Temp. Meter")]
        Temp_Meter,

        [System.ComponentModel.Description("Data Logger")]
        Data_Logger,

        [System.ComponentModel.Description("Power Meter")]
        Power_Meter,

        [System.ComponentModel.Description("USB")]
        USB,

        [System.ComponentModel.Description("Ethernet")]
        Ethernet,

        [System.ComponentModel.Description("Charge / Discharge Tester")]
        Charge_Discharge_Tester,

        [System.ComponentModel.Description("PD Emulator")]
        PD_Emulator,

        [System.ComponentModel.Description("BP Tester")]
        BP_Tester,

        [System.ComponentModel.Description("Thermal Camera")]
        Thermal_Camera,

        [System.ComponentModel.Description("DAQ Card")]
        DAQ_Card,

        [System.ComponentModel.Description("Battery")]
        Battery,

        [System.ComponentModel.Description("Battery Simulator")]
        Battery_Simulator,

        [System.ComponentModel.Description("LED Analyzer")]
        LED_Analyzer,

        [System.ComponentModel.Description("Rapid Charger")]
        Rapid_Charger,

        [System.ComponentModel.Description("EMF/EMI Analyzer")]
        EMF_EMI_Analyzer,

        [System.ComponentModel.Description("Electrical Safety Tester")]
        Electrical_Safety_Tester,

        [System.ComponentModel.Description("EV Gate")]
        EV_Gate,

        [System.ComponentModel.Description("LCR Meter")]
        LCR_Meter,


        // 第三組
        [System.ComponentModel.Description("Scanner")]
        Scanner,

        [System.ComponentModel.Description("RLC Load")]
        RLC_Load,

        [System.ComponentModel.Description("EVSE Signal Emulator")]
        EVSE_Signal_Emulator,

        [System.ComponentModel.Description("Low Voltage Isolation Box")]
        Low_Voltage_Isolation_Box,

        [System.ComponentModel.Description("Chamber")]
        Chamber,

        [System.ComponentModel.Description("SNMP")]
        SNMP,

        [System.ComponentModel.Description("EV CCS")]
        EV_CCS,


        // 類別與其他工具類
        [System.ComponentModel.Description("Arithmetic")]
        Arithmetic,

        [System.ComponentModel.Description("System, Flow Control")]
        System_Flow_Control,

        [System.ComponentModel.Description("String")]
        String,

        [System.ComponentModel.Description("Hold On Adjust")]
        Hold_On_Adjust,

        [System.ComponentModel.Description("Miscellaneous")]
        Miscellaneous,

        [System.ComponentModel.Description("HexString")]
        HexString,

        [System.ComponentModel.Description("Python")]
        Python,

        [System.ComponentModel.Description("Dll")]
        Dll,

        [System.ComponentModel.Description("Sub TI")]
        Sub_TI,

        [System.ComponentModel.Description("Thread TI")]
        Thread_TI
    }
}
