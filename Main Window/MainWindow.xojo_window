#tag Window
Begin Window MainWindow
   BackColor       =   &cFFFFFF00
   Backdrop        =   0
   CloseButton     =   True
   Compatibility   =   ""
   Composite       =   False
   Frame           =   0
   FullScreen      =   False
   FullScreenButton=   False
   HasBackColor    =   False
   Height          =   409
   ImplicitInstance=   True
   LiveResize      =   True
   MacProcID       =   0
   MaxHeight       =   32000
   MaximizeButton  =   False
   MaxWidth        =   610
   MenuBar         =   1855656149
   MenuBarVisible  =   True
   MinHeight       =   241
   MinimizeButton  =   True
   MinWidth        =   610
   Placement       =   0
   Resizeable      =   True
   Title           =   "Xojo Subnet Calculator Project"
   Visible         =   False
   Width           =   610
   Begin Label CurrentDisplayCounts_DataLabel
      AutoDeactivate  =   True
      Bold            =   False
      DataField       =   ""
      DataSource      =   ""
      Enabled         =   True
      Height          =   20
      HelpTag         =   "Tip: Scroll through this list by using your mouse wheel or mouse gestures"
      Index           =   -2147483648
      InitialParent   =   ""
      Italic          =   False
      Left            =   64
      LockBottom      =   True
      LockedInPosition=   False
      LockLeft        =   False
      LockRight       =   True
      LockTop         =   False
      Multiline       =   False
      Scope           =   0
      Selectable      =   False
      TabIndex        =   26
      TabPanelIndex   =   0
      Text            =   ""
      TextAlign       =   0
      TextColor       =   &c00000000
      TextFont        =   "Helvetica"
      TextSize        =   10.0
      TextUnit        =   0
      Top             =   345
      Transparent     =   True
      Underline       =   False
      Visible         =   False
      Width           =   133
   End
   Begin Separator Separator1
      AutoDeactivate  =   True
      Enabled         =   True
      Height          =   4
      HelpTag         =   ""
      Index           =   -2147483648
      InitialParent   =   ""
      Left            =   0
      LockBottom      =   False
      LockedInPosition=   False
      LockLeft        =   True
      LockRight       =   True
      LockTop         =   True
      Scope           =   0
      TabIndex        =   27
      TabPanelIndex   =   0
      TabStop         =   True
      Top             =   147
      Visible         =   True
      Width           =   610
   End
   Begin Calc_AllRanges Calc_AllRanges1
      AcceptFocus     =   False
      AcceptTabs      =   True
      AutoDeactivate  =   True
      BackColor       =   &cFFFFFF00
      Backdrop        =   0
      BeginOnClassLessSubnet=   False
      Enabled         =   True
      EraseBackground =   True
      HasBackColor    =   False
      Height          =   148
      HelpTag         =   ""
      InitialParent   =   ""
      Left            =   0
      LockBottom      =   False
      LockedInPosition=   False
      LockLeft        =   True
      LockRight       =   True
      LockTop         =   True
      Scope           =   0
      TabIndex        =   28
      TabPanelIndex   =   0
      TabStop         =   True
      Top             =   0
      Transparent     =   True
      UseFocusRing    =   False
      User_UseSingleRangeOnly=   False
      Visible         =   True
      Width           =   610
      Begin Separator Separator3
         AutoDeactivate  =   True
         Enabled         =   True
         Height          =   4
         HelpTag         =   ""
         Index           =   -2147483648
         InitialParent   =   "Calc_AllRanges1"
         Left            =   0
         LockBottom      =   False
         LockedInPosition=   False
         LockLeft        =   True
         LockRight       =   True
         LockTop         =   True
         Scope           =   0
         TabIndex        =   0
         TabPanelIndex   =   0
         TabStop         =   True
         Top             =   0
         Visible         =   True
         Width           =   610
      End
   End
   Begin Timer LazyLoadTimer
      Height          =   32
      Index           =   -2147483648
      InitialParent   =   ""
      Left            =   0
      LockedInPosition=   False
      Mode            =   0
      Period          =   1000
      Scope           =   0
      TabPanelIndex   =   0
      Top             =   0
      Width           =   32
   End
   Begin Listbox SubnetRangeListbox
      AutoDeactivate  =   True
      AutoHideScrollbars=   True
      Bold            =   False
      Border          =   True
      ColumnCount     =   1
      ColumnsResizable=   False
      ColumnWidths    =   ""
      DataField       =   ""
      DataSource      =   ""
      DefaultRowHeight=   -1
      Enabled         =   True
      EnableDrag      =   True
      EnableDragReorder=   False
      GridLinesHorizontal=   0
      GridLinesVertical=   0
      HasHeading      =   False
      HeadingIndex    =   -1
      Height          =   200
      HelpTag         =   ""
      Hierarchical    =   False
      Index           =   -2147483648
      InitialParent   =   ""
      InitialValue    =   ""
      Italic          =   False
      Left            =   -1
      LockBottom      =   True
      LockedInPosition=   False
      LockLeft        =   True
      LockRight       =   True
      LockTop         =   True
      RequiresSelection=   False
      Scope           =   0
      ScrollbarHorizontal=   False
      ScrollBarVertical=   True
      SelectionType   =   0
      TabIndex        =   31
      TabPanelIndex   =   0
      TabStop         =   False
      TextFont        =   "System"
      TextSize        =   0.0
      TextUnit        =   0
      Top             =   147
      Underline       =   False
      UseFocusRing    =   False
      Visible         =   True
      Width           =   611
      _ScrollOffset   =   0
      _ScrollWidth    =   -1
      Begin ScrollBar LazyLoadScrollBar
         AcceptFocus     =   True
         AutoDeactivate  =   True
         Enabled         =   True
         Height          =   180
         HelpTag         =   ""
         Index           =   -2147483648
         InitialParent   =   "SubnetRangeListbox"
         Left            =   594
         LineStep        =   1
         LiveScroll      =   False
         LockBottom      =   True
         LockedInPosition=   False
         LockLeft        =   False
         LockRight       =   True
         LockTop         =   True
         Maximum         =   100
         Minimum         =   0
         PageStep        =   20
         Scope           =   0
         TabIndex        =   0
         TabPanelIndex   =   0
         TabStop         =   True
         Top             =   165
         Value           =   0
         Visible         =   True
         Width           =   15
      End
   End
End
#tag EndWindow

#tag WindowCode
	#tag Event
		Sub Close()
		  quit(0)
		End Sub
	#tag EndEvent

	#tag Event
		Sub Minimize()
		  #if TargetMacOS Then
		    declare sub miniaturize lib "Cocoa" selector "miniaturize:" (windowRef as integer, id as Ptr)
		    miniaturize(self.handle, nil)
		  #endif
		End Sub
	#tag EndEvent

	#tag Event
		Sub Open()
		  // Open Window Position
		  Me.Left = 10
		  Me.Top = 50
		  
		  
		End Sub
	#tag EndEvent

	#tag Event
		Sub Resized()
		  #IF TargetWin32 Then
		    Separator1.Visible = True
		    Separator2.Visible = True
		  #ENDIF
		End Sub
	#tag EndEvent

	#tag Event
		Sub Resizing()
		  #IF TargetWin32 Then
		    // Hide Separators on Windows because it flickers bad
		    Separator1.Visible = False
		    Separator2.Visible = False
		  #ENDIF
		  
		  if MainWindow.Calc_AllRanges1.User_UseSingleRangeOnly = False Then
		    MainWindow.mUpdateDisplayRecordCounter
		  end if
		End Sub
	#tag EndEvent


	#tag Method, Flags = &h21
		Private Sub ConstructContextMenu(Base as MenuItem)
		  ConstructContextualReturnFlag = True
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub CreateArray()
		  if Calc_AllRanges1.User_UseSingleRangeOnly = False Then
		    
		    dim rangeInfoArray(-1, -1) as UInt32 = SubnetCalculatorSubClass1.Array_RangeInfo
		    dim decRangeInfoArray(-1, -1) as string = SubnetCalculatorSubClass1.Array_DecRangeInfo
		    
		    Dim y as Integer
		    Dim NetworkID, FirstIP,LastIP,BroadCastIP as String
		    dim lastSubnetIndex as integer = SubnetCalculatorSubClass1.Subnets - 1
		    for y = 0 to lastSubnetIndex
		      
		      // Network Subnet
		      rangeInfoArray(y,0) = SubnetCalculatorSubClass1.NetworkSubnetID
		      decRangeInfoArray(y,0) =  SubnetCalculatorSubClass1.fConvert_32BitDecimalTo8Bit_IP(rangeInfoArray(y,0))
		      
		      // First Host IP
		      SubnetCalculatorSubClass1.HostFirst_32BitDecimalWord = SubnetCalculatorSubClass1.NetworkSubnetID+1
		      rangeInfoArray(y,1) =SubnetCalculatorSubClass1.HostFirst_32BitDecimalWord
		      decRangeInfoArray(y,1) =SubnetCalculatorSubClass1.fConvert_32BitDecimalTo8Bit_IP(rangeInfoArray(y,1))
		      
		      // Reverse Mask
		      SubnetCalculatorSubClass1.ReverseSubnetMask = Bitwise.OnesComplement(SubnetCalculatorSubClass1.SubnetMask_32BitDecimalWord)
		      SubnetCalculatorSubClass1.BroadCastID_32BitDecimalWord = SubnetCalculatorSubClass1.NetworkSubnetID+SubnetCalculatorSubClass1.ReverseSubnetMask
		      
		      // BroadCast IP
		      rangeInfoArray(y,3) = SubnetCalculatorSubClass1.BroadCastID_32BitDecimalWord
		      SubnetCalculatorSubClass1.HostLast_32BitDecimalWord  = SubnetCalculatorSubClass1.BroadCastID_32BitDecimalWord-1
		      decRangeInfoArray(y,3) =SubnetCalculatorSubClass1.fConvert_32BitDecimalTo8Bit_IP(rangeInfoArray(y,3))
		      
		      // Last Host IP
		      rangeInfoArray(y,2) =SubnetCalculatorSubClass1.HostLast_32BitDecimalWord
		      decRangeInfoArray(y,2) = SubnetCalculatorSubClass1.fConvert_32BitDecimalTo8Bit_IP(rangeInfoArray(y,2))
		      
		      // Increment
		      SubnetCalculatorSubClass1.NetworkSubnetID = SubnetCalculatorSubClass1.BroadCastID_32BitDecimalWord+1
		      
		    Next y
		    
		  end if
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Sub DoContextAction(hititem as MenuItem)
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function fCalculateBinaryPadding(BinaryConversion as String) As String
		  
		  // Now Determine if we will need to Pad the Binary String with leading 0's since Xojo's Bin Function strips all Leading 0's
		  Dim BinaryConversion_Len as Integer = BinaryConversion.Len
		  
		  Select Case BinaryConversion_Len
		  Case 24
		    BinaryConversion = "00000000"+BinaryConversion
		  Case 25
		    BinaryConversion = "0000000"+BinaryConversion
		  Case 26
		    BinaryConversion = "000000"+BinaryConversion
		  Case 27
		    BinaryConversion = "00000"+BinaryConversion
		  Case 28
		    BinaryConversion = "0000"+BinaryConversion
		  Case 29
		    BinaryConversion = "000"+BinaryConversion
		  Case 30
		    BinaryConversion = "00"+BinaryConversion
		  Case 31
		    BinaryConversion = "0"+BinaryConversion
		  End Select
		  
		  Return BinaryConversion
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function fConvert2Bin(DecimalInput as String) As String
		  Dim FinalBinaryFormatted, Bin1, Bin2, Bin3, Bin4  as String
		  
		  Bin1 = DecimalInput.Left(8)
		  Bin2 = DecimalInput.Mid(9,8)
		  Bin3 = DecimalInput.Mid(17,8)
		  Bin4 = DecimalInput.Mid(25,8)
		  FinalBinaryFormatted = Bin1 + "." + Bin2  + "." + Bin3 + "." + Bin4
		  
		  Return FinalBinaryFormatted
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function fConvert2Hex(DecimalInput as String) As String
		  Dim HexFormatted, Hex1, Hex2, Hex3, Hex4  as String
		  
		  Hex1 = DecimalInput.Left(2)
		  Hex2 = DecimalInput.Mid(3,2)
		  Hex3 = DecimalInput.Mid(5,2)
		  Hex4 = DecimalInput.Mid(7,2)
		  HexFormatted = Hex1 + "-"+Hex2+"-"+Hex3+"-"+Hex4
		  
		  Return HexFormatted
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function fConvertNotationtoDecimal(inboundNotation as String) As String
		  Dim ConvertedValue as String
		  
		  Select Case inboundNotation
		  Case "1"
		    ConvertedValue = "128.0.0.0"
		  Case "2"
		    ConvertedValue = "192.0.0.0"
		  Case "3"
		    ConvertedValue = "224.0.0.0"
		  Case "4"
		    ConvertedValue = "240.0.0.0"
		  Case "5"
		    ConvertedValue = "248.0.0.0"
		  Case "6"
		    ConvertedValue = "252.0.0.0"
		  Case "7"
		    ConvertedValue = "254.0.0.0"
		  Case "8"
		    ConvertedValue = "255.0.0.0"
		  Case "9"
		    ConvertedValue = "255.128.0.0"
		  Case "10"
		    ConvertedValue = "255.192.0.0"
		  Case "11"
		    ConvertedValue = "255.224.0.0"
		  Case "12"
		    ConvertedValue = "255.240.0.0"
		  Case "13"
		    ConvertedValue = "255.248.0.0"
		  Case "14"
		    ConvertedValue = "255.252.0.0"
		  Case "15"
		    ConvertedValue = "255.254.0.0"
		  Case "16"
		    ConvertedValue = "255.255.0.0"
		  Case "17"
		    ConvertedValue = "255.255.128.0"
		  Case "18"
		    ConvertedValue = "255.255.192.0"
		  Case "19"
		    ConvertedValue = "255.255.224.0"
		  Case "20"
		    ConvertedValue = "255.255.240.0"
		  Case "21"
		    ConvertedValue = "255.255.248.0"
		  Case "22"
		    ConvertedValue = "255.255.252.0"
		  Case "23"
		    ConvertedValue = "255.255.254.0"
		  Case "24"
		    ConvertedValue = "255.255.255.0"
		  Case "25"
		    ConvertedValue = "255.255.255.128"
		  Case "26"
		    ConvertedValue = "255.255.255.192"
		  Case "27"
		    ConvertedValue = "255.255.255.224"
		  Case "28"
		    ConvertedValue = "255.255.255.240"
		  Case "29"
		    ConvertedValue = "255.255.255.248"
		  Case "30"
		    ConvertedValue = "255.255.255.252"
		  Case "31"
		    ConvertedValue = "255.255.255.254"
		  Case "32"
		    ConvertedValue = "255.255.255.255"
		  End Select
		  
		  Return ConvertedValue
		  
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function fConvert_32BitDecimalTo8Bit_IP(Inbound_32BitWordDecimal as UInt32) As String
		  // Convert 32 bit decimal base10 back to IP Address 8 bit Decimal
		  Dim Calc1, Octet1, Octet2, Octet3, Octet4, Base10IP As UInt32
		  Dim DecimalAddressParts(), DecimalAddress As String
		  Base10IP = Inbound_32BitWordDecimal
		  
		  For i As Integer = 0 To 3
		    Calc1 = Base10IP / 256^(3-i)
		    Base10IP = Base10IP - Calc1*(256^(3-i))
		    if i = 0 Then
		      Octet1 = Calc1
		    Elseif i = 1 Then
		      Octet2 = Calc1
		    Elseif i = 2 Then
		      Octet3 = Calc1
		    Elseif i = 3 Then
		      Octet4 = Calc1
		    End if
		  Next i
		  
		  DecimalAddressParts=Array(Str(Octet1),Str(Octet2),Str(Octet3), Str(Octet4))
		  DecimalAddress =Join(DecimalAddressParts,".")
		  
		  Return DecimalAddress
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function fIPv4_Validation(inputIPv4address as string) As String
		  
		  // First Measure How many periods we have (Should only have 4
		  Dim NumberOfPerids_RegEx as RegEx
		  Dim NumberOfPerids_RegExMatch as RegExMatch
		  Dim NumberOfPerids_HitText as String
		  Dim Counter as Integer = 0
		  // First Octet RegEx Parameters
		  NumberOfPerids_RegEx = New RegEx
		  NumberOfPerids_RegEx.Options.CaseSensitive = False
		  NumberOfPerids_RegEx.Options.Greedy = True
		  NumberOfPerids_RegEx.Options.MatchEmpty = True
		  NumberOfPerids_RegEx.Options.DotMatchAll = True
		  NumberOfPerids_RegEx.Options.StringBeginIsLineBegin = True
		  NumberOfPerids_RegEx.Options.StringEndIsLineEnd = True
		  NumberOfPerids_RegEx.Options.TreatTargetAsOneLine = False
		  NumberOfPerids_RegEx.SearchPattern = "[.]"
		  NumberOfPerids_RegExMatch = NumberOfPerids_RegEx.Search(inputIPv4address)
		  
		  While NumberOfPerids_RegExMatch <> Nil
		    NumberOfPerids_HitText = NumberOfPerids_RegExMatch.SubExpressionString(0)
		    NumberOfPerids_RegExMatch = NumberOfPerids_RegEx.Search
		    Counter = Counter + 1
		  Wend
		  
		  If Counter > 3 Then
		    // Too many periods
		    Return "Not_Valid"
		  Elseif Counter < 3 Then
		    // Not Enough periods
		    Return "Not_Valid"
		  Elseif Counter = 3 Then
		    // Good
		  end if
		  
		  
		  
		  // Break Down 4 Octets
		  Dim Octet_1 as String = inputIPv4address.NthField(".",1)
		  Dim Octet_2 as String = inputIPv4address.NthField(".",2)
		  Dim Octet_3 as String = inputIPv4address.NthField(".",3)
		  Dim Octet_4 as String = inputIPv4address.NthField(".",4)
		  
		  
		  
		  Dim Validate_Octet_1_RegEx as RegEx
		  Dim Validate_Octet_1_RegExMatch as RegExMatch
		  Dim Validate_Octet_1_HitText as String
		  Dim Validate_Octet_2_RegEx as RegEx
		  Dim Validate_Octet_2_RegExMatch as RegExMatch
		  Dim Validate_Octet_2_HitText as String
		  Dim Validate_Octet_3_RegEx as RegEx
		  Dim Validate_Octet_3_RegExMatch as RegExMatch
		  Dim Validate_Octet_3_HitText as String
		  Dim Validate_Octet_4_RegEx as RegEx
		  Dim Validate_Octet_4_RegExMatch as RegExMatch
		  Dim Validate_Octet_4_HitText as String
		  
		  // First Octet RegEx Parameters
		  Validate_Octet_1_RegEx = New RegEx
		  Validate_Octet_1_RegEx.Options.CaseSensitive = False
		  Validate_Octet_1_RegEx.Options.Greedy = True
		  Validate_Octet_1_RegEx.Options.MatchEmpty = True
		  Validate_Octet_1_RegEx.Options.DotMatchAll = True
		  Validate_Octet_1_RegEx.Options.StringBeginIsLineBegin = True
		  Validate_Octet_1_RegEx.Options.StringEndIsLineEnd = True
		  Validate_Octet_1_RegEx.Options.TreatTargetAsOneLine = False
		  Validate_Octet_1_RegEx.SearchPattern = "\b([1-9]|[1-9][0-9]|1[0-9][0-9]|2[0-1][0-9]|22[0-3])\b"
		  // Second Octet RegEx Parameters (allow 0-255)
		  Validate_Octet_2_RegEx = New RegEx
		  Validate_Octet_2_RegEx.Options.CaseSensitive = False
		  Validate_Octet_2_RegEx.Options.Greedy = True
		  Validate_Octet_2_RegEx.Options.MatchEmpty = True
		  Validate_Octet_2_RegEx.Options.DotMatchAll = True
		  Validate_Octet_2_RegEx.Options.StringBeginIsLineBegin = True
		  Validate_Octet_2_RegEx.Options.StringEndIsLineEnd = True
		  Validate_Octet_2_RegEx.Options.TreatTargetAsOneLine = False
		  Validate_Octet_2_RegEx.SearchPattern = "\b([0-9]|[1-9][0-9]|1[0-9][0-9]|2[0-4][0-9]|25[0-5])\b"
		  // Third Octet RegEx Parameters (allow 0-255)
		  Validate_Octet_3_RegEx = New RegEx
		  Validate_Octet_3_RegEx.Options.CaseSensitive = False
		  Validate_Octet_3_RegEx.Options.Greedy = True
		  Validate_Octet_3_RegEx.Options.MatchEmpty = True
		  Validate_Octet_3_RegEx.Options.DotMatchAll = True
		  Validate_Octet_3_RegEx.Options.StringBeginIsLineBegin = True
		  Validate_Octet_3_RegEx.Options.StringEndIsLineEnd = True
		  Validate_Octet_3_RegEx.Options.TreatTargetAsOneLine = False
		  Validate_Octet_3_RegEx.SearchPattern = "\b([0-9]|[1-9][0-9]|1[0-9][0-9]|2[0-4][0-9]|25[0-5])\b"
		  // Fourth Octet RegEx Parameters (allow 1-254)
		  Validate_Octet_4_RegEx = New RegEx
		  Validate_Octet_4_RegEx.Options.CaseSensitive = False
		  Validate_Octet_4_RegEx.Options.Greedy = True
		  Validate_Octet_4_RegEx.Options.MatchEmpty = True
		  Validate_Octet_4_RegEx.Options.DotMatchAll = True
		  Validate_Octet_4_RegEx.Options.StringBeginIsLineBegin = True
		  Validate_Octet_4_RegEx.Options.StringEndIsLineEnd = True
		  Validate_Octet_4_RegEx.Options.TreatTargetAsOneLine = False
		  Validate_Octet_4_RegEx.SearchPattern = "\b([0-9]|[1-9][0-9]|1[0-9][0-9]|2[0-1][0-9]|2[0-4][0-9]|25[0-4])\b"
		  
		  
		  // RegEx Process First Octet Results
		  Validate_Octet_1_RegExMatch = Validate_Octet_1_RegEx.Search(Octet_1)
		  if Validate_Octet_1_RegExMatch <> nil then
		    Validate_Octet_1_HitText = Validate_Octet_1_RegExMatch.SubExpressionString(0)
		    Validate_Octet_1_HitText = "Valid"
		    // Passed 1st Octet Validation - Move on to 2nd Octet Validation
		  elseif Validate_Octet_1_RegExMatch = nil then
		    // We did not have a Valid 1st Octet - Return
		    Return "Not_Valid"
		  end if
		  
		  // RegEx Process Second Octet Results
		  Validate_Octet_2_RegExMatch = Validate_Octet_2_RegEx.Search(Octet_2)
		  if Validate_Octet_2_RegExMatch <> nil then
		    Validate_Octet_2_HitText = Validate_Octet_2_RegExMatch.SubExpressionString(0)
		    Validate_Octet_2_HitText = "Valid"
		    // Passed 2nd Octet Validation - Move on to 3rd Octet Validation
		  elseif Validate_Octet_2_RegExMatch = nil then
		    // We did not have a Valid 2nd Octet - Return
		    Return "Not_Valid"
		  end if
		  
		  // RegEx Process Third Octet Results
		  Validate_Octet_3_RegExMatch = Validate_Octet_3_RegEx.Search(Octet_3)
		  if Validate_Octet_3_RegExMatch <> nil then
		    Validate_Octet_3_HitText = Validate_Octet_3_RegExMatch.SubExpressionString(0)
		    Validate_Octet_3_HitText = "Valid"
		    // Passed 3rd Octet Validation - Move on to 4th Octet Validation
		  elseif Validate_Octet_3_RegExMatch = nil then
		    // We did not have a Valid 3rd Octet - Return
		    Return "Not_Valid"
		  end if
		  
		  // RegEx Process Fourth and FINAL Octet
		  Validate_Octet_4_RegExMatch = Validate_Octet_4_RegEx.Search(Octet_4)
		  if Validate_Octet_4_RegExMatch <> nil then
		    Validate_Octet_4_HitText = Validate_Octet_4_RegExMatch.SubExpressionString(0)
		    Validate_Octet_4_HitText = "Valid"
		    // Passed 4th Octet Validation - Return Successful
		    Return "Valid"
		  elseif Validate_Octet_4_RegExMatch = nil then
		    // We did not have a Valid 3rd Octet - Return
		    Return "Not_Valid"
		  end if
		  
		  
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function fSubnetMaskValidation(inputMask as string) As String
		  // Did we receive a /xx Notation? IF so lets convert it
		  Dim DetectShortFormNotation_RegEx as RegEx
		  Dim DetectShortFormNotation_RegExMatch as RegExMatch
		  Dim DetectShortFormNotation_HitText as String
		  Dim CalculateNotation2Decimal as String
		  DetectShortFormNotation_RegEx = New RegEx
		  DetectShortFormNotation_RegEx.Options.CaseSensitive = False
		  DetectShortFormNotation_RegEx.Options.Greedy = True
		  DetectShortFormNotation_RegEx.Options.MatchEmpty = True
		  DetectShortFormNotation_RegEx.Options.DotMatchAll = True
		  DetectShortFormNotation_RegEx.Options.StringBeginIsLineBegin = True
		  DetectShortFormNotation_RegEx.Options.StringEndIsLineEnd = True
		  DetectShortFormNotation_RegEx.Options.TreatTargetAsOneLine = False
		  DetectShortFormNotation_RegEx.SearchPattern = "(^/)\b([8-9]|1[0-9]|2[0-9]|3[0-2])\b"
		  
		  // RegEx Process Short Form Notation
		  DetectShortFormNotation_RegExMatch = DetectShortFormNotation_RegEx.Search(inputMask)
		  if DetectShortFormNotation_RegExMatch <> nil then
		    // Detected ShortForm Notation
		    DetectShortFormNotation_HitText = DetectShortFormNotation_RegExMatch.SubExpressionString(0)
		    
		    // Strip off the /
		    Dim Notation as String = Trim(DetectShortFormNotation_HitText.Mid(2,2))
		    // Convert Notation into Decimal
		    CalculateNotation2Decimal = fConvertNotationtoDecimal(Notation)
		    inputMask = CalculateNotation2Decimal
		  elseif DetectShortFormNotation_RegExMatch = nil then
		    // Did not Detect Short Form Notation .. Continue on to Decimal Based Subnet Mask Validation
		  end if
		  
		  // Decimal Based Subnet Validation
		  Dim Validate_SubnetMask_RegEx as RegEx
		  Dim Validate_SubnetMask_RegExMatch as RegExMatch
		  Dim Validate_SubnetMask_HitText as String
		  
		  Validate_SubnetMask_RegEx = New RegEx
		  Validate_SubnetMask_RegEx.Options.CaseSensitive = False
		  Validate_SubnetMask_RegEx.Options.Greedy = True
		  Validate_SubnetMask_RegEx.Options.MatchEmpty = True
		  Validate_SubnetMask_RegEx.Options.DotMatchAll = True
		  Validate_SubnetMask_RegEx.Options.StringBeginIsLineBegin = True
		  Validate_SubnetMask_RegEx.Options.StringEndIsLineEnd = True
		  Validate_SubnetMask_RegEx.Options.TreatTargetAsOneLine = False
		  Validate_SubnetMask_RegEx.SearchPattern = "\b(128.0.0.0|192.0.0.0|224.0.0.0|240.0.0.0|248.0.0.0|252.0.0.0|254.0.0.0|255.0.0.0|255.128.0.0|255.192.0.0|255.224.0.0|255.240.0.0|255.248.0.0|255.252.0.0|255.254.0.0|255.255.0.0|255.255.128.0|255.255.192.0|255.255.224.0|255.255.240.0|255.255.248.0|255.255.252.0|255.255.254.0|255.255.255.0|255.255.255.128|255.255.255.192|255.255.255.224|255.255.255.240|255.255.255.248|255.255.255.252|255.255.255.254)\b"
		  
		  // RegEx Process Subnet Mask Validation
		  Validate_SubnetMask_RegExMatch = Validate_SubnetMask_RegEx.Search(inputMask)
		  if Validate_SubnetMask_RegExMatch <> nil then
		    Validate_SubnetMask_HitText = Validate_SubnetMask_RegExMatch.SubExpressionString(0)
		    
		    Return Validate_SubnetMask_HitText
		  elseif Validate_SubnetMask_RegExMatch = nil then
		    Return "Not_Valid"
		  end if
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub mCaptureSelectedRows()
		  Dim i as Integer
		  For i =0 to SubnetRangeListbox.ListCount - 1
		    if SubnetRangeListbox.Selected(i) = True then
		      MainListboxSelected.Append i
		    end if
		  Next
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Sub mClear()
		  // Delete all Rows in Listbox
		  SubnetRangeListbox.DeleteAllRows
		  
		  // Clear Other TextFields
		  Calc_AllRanges1.TotalNumOfHosts_Data.Text = ""
		  Calc_AllRanges1.TotalNumOfSubnets_Data.Text = ""
		  
		  Calc_AllRanges1.ValidateAllRange_IP_TextField.Text = ""
		  Calc_AllRanges1.ValidateAllRange_SubnetMask_TextField.Text = ""
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub mUpdateDisplayRecordCounter()
		  // Initial Listbox Counter Message on UI
		  MainWindow.CurrentDisplayCounts_DataLabel.Visible = True
		  
		  CalcRowHeightIncrease  = 7
		  if MainWindow.Height > 409 Then
		    CalcRowHeightIncrease = Ceil(CalcRowHeightIncrease+(MainWindow.Height - 409) / 24)
		  end if
		  
		  if MainWindow.Height > 490 OR MainWindow.SubnetRangeListbox.ListCount >= 11 Then
		    if MainWindow.SubnetRangeListbox.ScrollPosition+CalcRowHeightIncrease > SubnetCalculatorSubClass1.Subnets Then Exit Sub
		    MainWindow.CurrentDisplayCounts_DataLabel.Text = Format((MainWindow.SubnetRangeListbox.ScrollPosition+CalcRowHeightIncrease),"###,###,###") + " of " + Format(SubnetCalculatorSubClass1.Subnets,"###,###,###")
		  Elseif SubnetRangeListbox.ListCount < 11 Then
		    MainWindow.CurrentDisplayCounts_DataLabel.Text = Format(SubnetCalculatorSubClass1.Subnets,"###,###,###") + " of " + Format(SubnetCalculatorSubClass1.Subnets,"###,###,###")
		  End If
		  
		  
		End Sub
	#tag EndMethod


	#tag Property, Flags = &h0
		CalcRowHeightIncrease As double
	#tag EndProperty

	#tag Property, Flags = &h21
		Private ConstructContextualReturnFlag As Boolean = False
	#tag EndProperty

	#tag Property, Flags = &h21
		Private ContextualReturnFlag As Boolean = False
	#tag EndProperty

	#tag Property, Flags = &h0
		EndTime As Integer
	#tag EndProperty

	#tag Property, Flags = &h0
		IsMaximized As Boolean
	#tag EndProperty

	#tag Property, Flags = &h0
		IsMinimized As Boolean
	#tag EndProperty

	#tag Property, Flags = &h21
		Private ListboxLastScrollValue As Integer
	#tag EndProperty

	#tag Property, Flags = &h0
		MainListboxSelected() As Integer
	#tag EndProperty

	#tag Property, Flags = &h0
		Queue() As String
	#tag EndProperty

	#tag Property, Flags = &h21
		Private SessionTimerCountdownValue As String
	#tag EndProperty

	#tag Property, Flags = &h0
		StartTime As Double
	#tag EndProperty

	#tag Property, Flags = &h0
		SubnetCalculatorSubClass1 As SubnetCalculator_Class
	#tag EndProperty


#tag EndWindowCode

#tag Events Separator1
	#tag Event
		Sub Open()
		  #IF TargetWin32 Then
		    Me.Top = 146
		  #ENDIF
		End Sub
	#tag EndEvent
#tag EndEvents
#tag Events Separator3
	#tag Event
		Sub Open()
		  #IF TargetMacOS Then
		    Me.Visible = False
		  #ELSEIF TargetWin32 Then
		    Me.Top = -2
		  #ENDIF
		End Sub
	#tag EndEvent
#tag EndEvents
#tag Events LazyLoadTimer
	#tag Event
		Sub Action()
		  
		  if SubnetRangeListbox.ListCount <= SubnetCalculatorSubClass1.Subnets-1  Then
		    SubnetCalculatorSubClass1.mLoadListbox(10000)
		  Else
		    Me.Mode = Timer.ModeOff
		  end if
		  
		  
		End Sub
	#tag EndEvent
#tag EndEvents
#tag Events SubnetRangeListbox
	#tag Event
		Sub Open()
		  //// First Instantation of our Subnet Calculator Class
		  //SubnetCalculatorSubClass1 = New SubnetCalculator_Class
		  
		  
		  
		  // Setup Default Behavior
		  Me.ColumnsResizable = True
		  Me.RequiresSelection = False
		  Me.ColumnCount = 5
		  Me.HasHeading = True
		  Me.AutoHideScrollbars = false
		  Me.ScrollBarHorizontal = False
		  Me.ScrollBarVertical = False
		  Me.UseFocusRing = False
		  Me.ColumnWidths = "66,112,117,118"
		  Me.TextFont = "Helvetica"
		  Me.TextSize = 12
		  Me.SelectionType = Listbox.SelectionMultiple
		  Me.DefaultRowHeight = 24
		  
		  Me.Heading(0) = Strings_Module.kMainWindow_Listbox_SubnetNum
		  Me.HeaderType(0) =listbox.HeaderTypes.Sortable
		  Me.Heading(1) = Strings_Module.kMainWindow_Listbox_NetworkID
		  Me.Heading(2) = Strings_Module.kMainWindow_Listbox_FirstHost
		  Me.Heading(3) = Strings_Module.kMainWindow_Listbox_LastHost
		  Me.Heading(4) = Strings_Module.kMainWindow_Listbox_BroadcastHost
		  
		  
		End Sub
	#tag EndEvent
	#tag Event
		Function CellBackgroundPaint(g As Graphics, row As Integer, column As Integer) As Boolean
		  // Get Cell Width
		  Dim RW as integer = me.Column(column).WidthActual
		  Dim RH as integer = me.RowHeight
		  
		  //Fill background Cell color (multicolor) of Only Added Rows
		  If row Mod 2=0 then
		    g.foreColor = App.ThemeColor_Listbox_Color1
		  else
		    g.foreColor = App.ThemeColor_Listbox_Color2
		  end if
		  g.FillRect 1,1,RW,RH
		  
		  if me.selected(Row) Then
		    g.ForeColor = &ceafbff
		    g.FillRect 0,0,RW,RH
		  End if
		  
		  //Draw the HighlightColor for the Cell Outline for New Row
		  if me.Selected(row) Then //AND Column = 0 OR me.Selected(row) AND Column = 1 OR me.Selected(row) AND Column = 2 OR me.Selected(row) AND Column = 3 OR me.Selected(row) AND Column = 4 Then
		    g.PenHeight= 1
		    g.PenWidth = 1
		    g.ForeColor = &cC0C0C0//&c5391F7
		    g.DrawRect(0, 0, RW,RH)
		    
		  Else
		    g.Transparency = 50
		    g.PenHeight= 1
		    g.PenWidth = 1
		    g.ForeColor = &c000000
		    g.DrawRect(0, 0, RW,RH)
		    g.Transparency = 0
		  End if
		  
		  //Add a bottom bar to even out the thickness of the last added bottom outline
		  if row = me.LastIndex+1 then
		    if me.Selected(row) Then //AND Column = 0 OR me.Selected(row) AND Column = 1 OR me.Selected(row) AND Column = 2 OR me.Selected(row) AND Column = 3 OR me.Selected(row) AND Column = 4 Then
		      g.PenHeight= 2
		      g.PenWidth =2
		      g.ForeColor =  &c5391F7
		      g.DrawLine(0, 0, RW,0)
		    Else
		      g.Transparency = 50
		      g.PenHeight= 1
		      g.PenWidth = 1
		      g.ForeColor = &c000000
		      g.DrawLine(0, 0, RW,0)
		      g.Transparency = 0
		    End if
		  end if
		  
		  Return True
		  
		  
		  
		End Function
	#tag EndEvent
	#tag Event
		Function CellTextPaint(g As Graphics, row As Integer, column As Integer, x as Integer, y as Integer) As Boolean
		  #pragma Unused x
		  #pragma Unused y
		  
		  // Set Default Font Parms
		  g.TextFont = "System"
		  g.ForeColor = &c000000 // Set Text for all cells to black
		  g.TextSize = 12
		End Function
	#tag EndEvent
	#tag Event
		Function MouseWheel(X As Integer, Y As Integer, deltaX as Integer, deltaY as Integer) As Boolean
		  #if not DebugBuild
		    #pragma BackgroundTasks false
		    #pragma BoundsChecking false
		    #pragma NilObjectChecking false
		    #pragma StackOverflowChecking false
		  #endif
		  
		  //Return True
		  // Dont allow scrolling when we only have less than 11 Entries
		  
		  if MainWindow.SubnetRangeListbox.ListCount <=11 then
		    Return True
		  End if
		  
		  IF Calc_AllRanges1.User_UseSingleRangeOnly = False Then
		    if Me.ListCount <> 0  Then
		      if UBound(SubnetCalculatorSubClass1.Array_RangeInfo) <> -1 Then
		        if SubnetCalculatorSubClass1.Subnets <> Me.ListCount Then
		          //SubnetCalculatorSubClass1.mLoadListbox(40)
		        Elseif Me.ListCount <= SubnetCalculatorSubClass1.Subnets-1 Then
		          Redim SubnetCalculatorSubClass1.Array_RangeInfo(-1,-1)
		        End if
		        
		      ELSE
		        Return True
		      End if
		    End if
		    
		  Elseif Calc_AllRanges1.User_UseSingleRangeOnly = True Then
		    // Prevent Scrolling
		    Return True
		  End if
		  
		  // Sync Mouse Wheel with Scroll Bar
		  if LazyLoadScrollBar.Visible = True Then
		    LazyLoadScrollBar.Value = Me.ScrollPosition+6
		  End if
		  
		  // Update Display Counters on Main Window
		  mUpdateDisplayRecordCounter
		  
		  
		  
		  
		End Function
	#tag EndEvent
	#tag Event
		Function ConstructContextualMenu(base as MenuItem, x as Integer, y as Integer) As Boolean
		  ConstructContextMenu(base)
		  
		  if ConstructContextualReturnFlag = True then
		    ConstructContextualReturnFlag = false
		    Return True
		  elseif ConstructContextualReturnFlag = False then
		    Return False
		  end if
		End Function
	#tag EndEvent
	#tag Event
		Function ContextualMenuAction(hitItem as MenuItem) As Boolean
		  // If the ListBox is Empty then just Return
		  if ContextualReturnFlag = True then
		    ContextualReturnFlag = false
		    Return True
		  elseif ContextualReturnFlag = False then
		    Return False
		  end if
		End Function
	#tag EndEvent
	#tag Event
		Function MouseDown(x As Integer, y As Integer) As Boolean
		  If  NOT IsContextualClick then Return False
		  
		  Dim base as New menuItem
		  ConstructContextMenu(base)
		  Dim hitItem as MenuItem =  Base.Popup
		  DoContextAction(hitItem)
		  
		  Return True
		End Function
	#tag EndEvent
	#tag Event
		Function CompareRows(row1 as Integer, row2 as Integer, column as Integer, ByRef result as Integer) As Boolean
		  If Val(Me.Cell(row1, 0)) < Val(Me.Cell(row2, 0)) Then
		    result = -1
		  ElseIf Val(Me.Cell(row1, 0)) > Val(Me.Cell(row2, 0)) Then
		    result = 1
		  Else
		    result = 0
		  End If
		  
		  Return True
		  
		End Function
	#tag EndEvent
#tag EndEvents
#tag Events LazyLoadScrollBar
	#tag Event
		Sub ValueChanged()
		  #if not DebugBuild
		    #pragma BackgroundTasks false
		    #pragma BoundsChecking false
		    #pragma NilObjectChecking false
		    #pragma StackOverflowChecking false
		  #endif
		  
		  // This aligns the scrollbar perfectly on the UBound Boundary
		  Dim NumberOfCells  as Integer = 6
		  Dim Rows as Integer
		  if MainWindow.Height > 409 Then
		    Rows =Ceil((MainWindow.Height - 409) / 24)
		    NumberOfCells = Rows + 6 -1
		  end if
		  
		  SubnetRangeListbox.ScrollPosition = Me.Value-NumberOfCells
		  
		  if SubnetCalculatorSubClass1.Subnets <> SubnetRangeListbox.ListCount Then
		    if SubnetCalculatorSubClass1.Subnets > 600000 Then
		      SubnetCalculatorSubClass1.mLoadListbox(10000)
		    Else
		      SubnetCalculatorSubClass1.mLoadListbox(SubnetCalculatorSubClass1.Subnets/6)
		    end if
		    
		    
		  Elseif SubnetRangeListbox.ListCount <= SubnetCalculatorSubClass1.Subnets-1 Then
		    Redim SubnetCalculatorSubClass1.Array_RangeInfo(-1,-1)
		  End if
		  
		  // Update Display Counters on Main Window
		  mUpdateDisplayRecordCounter
		End Sub
	#tag EndEvent
	#tag Event
		Sub Open()
		  Me.LiveScroll = True
		  Me.LineStep = 1
		  
		  
		End Sub
	#tag EndEvent
#tag EndEvents
#tag ViewBehavior
	#tag ViewProperty
		Name="BackColor"
		Visible=true
		Group="Appearance"
		InitialValue="&hFFFFFF"
		Type="Color"
	#tag EndViewProperty
	#tag ViewProperty
		Name="Backdrop"
		Visible=true
		Group="Appearance"
		Type="Picture"
		EditorType="Picture"
	#tag EndViewProperty
	#tag ViewProperty
		Name="CalcRowHeightIncrease"
		Group="Behavior"
		Type="double"
	#tag EndViewProperty
	#tag ViewProperty
		Name="CloseButton"
		Visible=true
		Group="Appearance"
		InitialValue="True"
		Type="Boolean"
		EditorType="Boolean"
	#tag EndViewProperty
	#tag ViewProperty
		Name="Composite"
		Visible=true
		Group="Appearance"
		InitialValue="False"
		Type="Boolean"
	#tag EndViewProperty
	#tag ViewProperty
		Name="EndTime"
		Group="Behavior"
		Type="Integer"
	#tag EndViewProperty
	#tag ViewProperty
		Name="Frame"
		Visible=true
		Group="Appearance"
		InitialValue="0"
		Type="Integer"
		EditorType="Enum"
		#tag EnumValues
			"0 - Document"
			"1 - Movable Modal"
			"2 - Modal Dialog"
			"3 - Floating Window"
			"4 - Plain Box"
			"5 - Shadowed Box"
			"6 - Rounded Window"
			"7 - Global Floating Window"
			"8 - Sheet Window"
			"9 - Metal Window"
			"10 - Drawer Window"
			"11 - Modeless Dialog"
		#tag EndEnumValues
	#tag EndViewProperty
	#tag ViewProperty
		Name="FullScreen"
		Group="Appearance"
		InitialValue="False"
		Type="Boolean"
		EditorType="Boolean"
	#tag EndViewProperty
	#tag ViewProperty
		Name="FullScreenButton"
		Visible=true
		Group="Appearance"
		InitialValue="False"
		Type="Boolean"
		EditorType="Boolean"
	#tag EndViewProperty
	#tag ViewProperty
		Name="HasBackColor"
		Visible=true
		Group="Appearance"
		InitialValue="False"
		Type="Boolean"
	#tag EndViewProperty
	#tag ViewProperty
		Name="Height"
		Visible=true
		Group="Position"
		InitialValue="400"
		Type="Integer"
	#tag EndViewProperty
	#tag ViewProperty
		Name="ImplicitInstance"
		Visible=true
		Group="Appearance"
		InitialValue="True"
		Type="Boolean"
		EditorType="Boolean"
	#tag EndViewProperty
	#tag ViewProperty
		Name="Interfaces"
		Visible=true
		Group="ID"
		Type="String"
	#tag EndViewProperty
	#tag ViewProperty
		Name="IsMaximized"
		Group="Behavior"
		Type="Boolean"
	#tag EndViewProperty
	#tag ViewProperty
		Name="IsMinimized"
		Group="Behavior"
		Type="Boolean"
	#tag EndViewProperty
	#tag ViewProperty
		Name="LiveResize"
		Visible=true
		Group="Appearance"
		InitialValue="True"
		Type="Boolean"
		EditorType="Boolean"
	#tag EndViewProperty
	#tag ViewProperty
		Name="MacProcID"
		Visible=true
		Group="Appearance"
		InitialValue="0"
		Type="Integer"
	#tag EndViewProperty
	#tag ViewProperty
		Name="MaxHeight"
		Visible=true
		Group="Position"
		InitialValue="32000"
		Type="Integer"
	#tag EndViewProperty
	#tag ViewProperty
		Name="MaximizeButton"
		Visible=true
		Group="Appearance"
		InitialValue="True"
		Type="Boolean"
		EditorType="Boolean"
	#tag EndViewProperty
	#tag ViewProperty
		Name="MaxWidth"
		Visible=true
		Group="Position"
		InitialValue="32000"
		Type="Integer"
	#tag EndViewProperty
	#tag ViewProperty
		Name="MenuBar"
		Visible=true
		Group="Appearance"
		Type="MenuBar"
		EditorType="MenuBar"
	#tag EndViewProperty
	#tag ViewProperty
		Name="MenuBarVisible"
		Group="Appearance"
		InitialValue="True"
		Type="Boolean"
		EditorType="Boolean"
	#tag EndViewProperty
	#tag ViewProperty
		Name="MinHeight"
		Visible=true
		Group="Position"
		InitialValue="64"
		Type="Integer"
	#tag EndViewProperty
	#tag ViewProperty
		Name="MinimizeButton"
		Visible=true
		Group="Appearance"
		InitialValue="True"
		Type="Boolean"
		EditorType="Boolean"
	#tag EndViewProperty
	#tag ViewProperty
		Name="MinWidth"
		Visible=true
		Group="Position"
		InitialValue="64"
		Type="Integer"
	#tag EndViewProperty
	#tag ViewProperty
		Name="Name"
		Visible=true
		Group="ID"
		Type="String"
	#tag EndViewProperty
	#tag ViewProperty
		Name="Placement"
		Visible=true
		Group="Position"
		InitialValue="0"
		Type="Integer"
		EditorType="Enum"
		#tag EnumValues
			"0 - Default"
			"1 - Parent Window"
			"2 - Main Screen"
			"3 - Parent Window Screen"
			"4 - Stagger"
		#tag EndEnumValues
	#tag EndViewProperty
	#tag ViewProperty
		Name="Resizeable"
		Visible=true
		Group="Appearance"
		InitialValue="True"
		Type="Boolean"
		EditorType="Boolean"
	#tag EndViewProperty
	#tag ViewProperty
		Name="StartTime"
		Group="Behavior"
		Type="Double"
	#tag EndViewProperty
	#tag ViewProperty
		Name="Super"
		Visible=true
		Group="ID"
		Type="String"
	#tag EndViewProperty
	#tag ViewProperty
		Name="Title"
		Visible=true
		Group="Appearance"
		InitialValue="Untitled"
		Type="String"
	#tag EndViewProperty
	#tag ViewProperty
		Name="Visible"
		Visible=true
		Group="Appearance"
		InitialValue="True"
		Type="Boolean"
		EditorType="Boolean"
	#tag EndViewProperty
	#tag ViewProperty
		Name="Width"
		Visible=true
		Group="Position"
		InitialValue="600"
		Type="Integer"
	#tag EndViewProperty
#tag EndViewBehavior
