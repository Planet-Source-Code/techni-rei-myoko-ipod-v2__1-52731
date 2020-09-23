Attribute VB_Name = "BatteryStatus"
Option Explicit

Private Type SYSTEM_POWER_STATUS
    ACLineStatus As Byte
    BatteryFlag As Byte
    BatteryLifePercent As Byte
    Reserved1 As Byte
    BatteryLifeTime As Long
    BatteryFullLifeTime As Long
End Type

Public Enum BatteryConstants
    Bat_High = 1
    Bat_Low = 2
    Bat_Critical = 4
    Bat_Charging = 8
    Bat_None = 128
    Bat_Unknown = 255
End Enum

Private Declare Function GetSystemPowerStatus Lib "kernel32" (lpSystemPowerStatus As SYSTEM_POWER_STATUS) As Long

Public Function BatteryPercent(Optional ByRef Flag As BatteryConstants) As Long
    Dim SPS As SYSTEM_POWER_STATUS
    GetSystemPowerStatus SPS
    BatteryPercent = SPS.BatteryLifePercent / 2.55
    Flag = SPS.BatteryFlag
End Function
Public Function BatteryFlag() As BatteryConstants
    Dim temp As BatteryConstants
    Call BatteryPercent(temp)
    BatteryFlag = temp
End Function
