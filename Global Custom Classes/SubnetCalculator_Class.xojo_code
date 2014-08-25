#tag Class
Protected Class SubnetCalculator_Class
	#tag Method, Flags = &h21
		Private Function fBuild_IPv4Formatted_String(Octet1 as Integer, Octet2 as Integer, Octet3 as Integer, Octet4 as integer) As String
		  dim DecimalAddressParts() as string=Array(Str(Octet1),Str(Octet2),Str(Octet3), Str(Octet4))
		  dim DecimalAddress as String =join(DecimalAddressParts,".")
		  
		  Return DecimalAddress
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function fCalculateHostBits() As Integer
		  Dim n2,n3,n4,m2,m3,m4 as Integer
		  
		  n2 = SubnetMask2Dec
		  Select Case n2
		  Case 0
		    m2 = 8
		  Case 128
		    m2 = 7
		  Case 192
		    m2 = 6
		  Case 224
		    m2 = 5
		  Case 240
		    m2 = 4
		  Case 248
		    m2 = 3
		  Case 252
		    m2 = 2
		  Case 254
		    m2 = 1
		  Case 255
		    m2 = 0
		  End Select
		  
		  n3 = SubnetMask3Dec
		  
		  Select Case n3
		  Case 0
		    m3 = 8
		  Case 128
		    m3 = 7
		  Case 192
		    m3 = 6
		  Case 224
		    m3 = 5
		  Case 240
		    m3 = 4
		  Case 248
		    m3 = 3
		  Case 252
		    m3 = 2
		  Case 254
		    m3 = 1
		  Case 255
		    m3 = 0
		  End Select
		  
		  n4 = SubnetMask4Dec
		  Select Case n4
		  Case 0
		    m4 = 8
		  Case 128
		    m4 = 7
		  Case 192
		    m4 = 6
		  Case 224
		    m4 = 5
		  Case 240
		    m4 = 4
		  Case 248
		    m4 = 3
		  Case 252
		    m4 = 2
		  Case 254
		    m4 = 1
		  Case 255
		    m4 = 0
		  End Select
		  
		  Return m2+m3+m4
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function fCalculateSubnetBits() As Integer
		  Dim n2,n3,n4,m2,m3,m4 as Integer
		  
		  n2 = SubnetMask2Dec
		  Select Case n2
		  Case 0
		    m2 = 0
		  Case 128
		    m2 = 1
		  Case 192
		    m2 = 2
		  Case 224
		    m2 = 3
		  Case 240
		    m2 = 4
		  Case 248
		    m2 = 5
		  Case 252
		    m2 = 6
		  Case 254
		    m2 = 7
		  Case 255
		    m2 = 8
		  End Select
		  
		  n3 = SubnetMask3Dec
		  
		  Select Case n3
		  Case 0
		    m3 = 0
		  Case 128
		    m3 = 1
		  Case 192
		    m3 = 2
		  Case 224
		    m3 = 3
		  Case 240
		    m3 = 4
		  Case 248
		    m3 = 5
		  Case 252
		    m3 = 6
		  Case 254
		    m3 = 7
		  Case 255
		    m3 = 8
		  End Select
		  
		  n4 = SubnetMask4Dec
		  Select Case n4
		  Case 0
		    m4 = 0
		  Case 128
		    m4 = 1
		  Case 192
		    m4 = 2
		  Case 224
		    m4 = 3
		  Case 240
		    m4 = 4
		  Case 248
		    m4 = 5
		  Case 252
		    m4 = 6
		  Case 254
		    m4 = 7
		  Case 255
		    m4 = 8
		  End Select
		  
		  SubnetPrefixLength = 8+m2+m3+m4
		  
		  Return m2+m3+m4
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function fConvert_32BitDecimalTo8Bit_IP(Inbound_32BitWordDecimal as UInt32) As String
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
		  DecimalAddress=Join(DecimalAddressParts,".")
		  
		  Return DecimalAddress
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function fConvert_8BitDecimalTo32BitDecimal(Octet1 as UInt32, Octet2 as UInt32, Octet3 as Uint32, Octet4 as UInt32) As UInt32
		  // Convert IP Address 8 bit Decimal to 32 bit decimal base10
		  dim Base10IP,a,b,c,d as UInt32
		  
		  a = Octet1 * 16777216
		  b = Octet2 * 65536
		  c =  Octet3 * 256
		  d =  Octet4
		  
		  Base10IP = a+b+c+d
		  
		  Return Base10IP
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function fGetClassFullNetwork(Input_StartIP_32BitDecimalWord as UInt32, Input_SubnetMask_32BitDecimalWord as uInt32) As UInt32
		  // Get the ClassFull Network Based on the Network ID Passed In - Not Comparing it to the Subnet Mask
		  Dim ClassFullNetworkID as UInt32
		  
		  if Input_StartIP_32BitDecimalWord >= 16777216 AND Input_StartIP_32BitDecimalWord <= 2130706431 Then
		    // Class A 1 - 126
		    // Compare our Network Against 255.0.0.0
		    IANA_Class = "Class A"
		    ClassFullNetworkID =  Input_StartIP_32BitDecimalWord AND 4278190080
		    Return ClassFullNetworkID
		  Else
		    // Class B 128 - 191
		    if Input_StartIP_32BitDecimalWord >= 2147483648 AND Input_StartIP_32BitDecimalWord <= 3221225471 Then
		      // Matched B
		      IANA_Class = "Class B"
		      ClassFullNetworkID =  Input_StartIP_32BitDecimalWord AND 4294901760
		      Return ClassFullNetworkID
		    else
		      // Class C 192-223
		      If Input_StartIP_32BitDecimalWord >= 3221225472 AND Input_StartIP_32BitDecimalWord <= 3758096383 Then
		        // Matched C
		        IANA_Class = "Class C"
		        ClassFullNetworkID =  Input_StartIP_32BitDecimalWord AND 4294967040
		        Return ClassFullNetworkID
		      Else
		        // D (Multicast) 224 - 239
		        If Input_StartIP_32BitDecimalWord >= 3758096384 AND Input_StartIP_32BitDecimalWord <= 4026531839 Then
		          // Matched D (Multicast)
		          IANA_Class = "Class D Multicast"
		          ClassFullNetworkID =  Input_StartIP_32BitDecimalWord AND 4294967295
		          Return ClassFullNetworkID
		        End If
		      End if
		      
		    end if
		    
		    
		  end if
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Sub fGetSingleRanges(Input_StartIP_32BitDecimalWord as UInt32, Input_SubnetMask_32BitDecimalWord as uInt32)
		  #pragma BackgroundTasks false
		  #pragma BoundsChecking false
		  #pragma NilObjectChecking false
		  #pragma StackOverflowChecking false
		  #pragma DisableBackgroundTasks
		  
		  Dim BroadcastID_32BitDecimalWord,NetworkSubnetID as UInt32
		  Dim HostFirst_32BitDecimalWord, HostLast_32BitDecimalWord as UInt32
		  
		  NetworkSubnetID = Input_StartIP_32BitDecimalWord AND Input_SubnetMask_32BitDecimalWord
		  // Determine ClassFul Class Type
		  mGetClassOnly(Input_StartIP_32BitDecimalWord,Input_SubnetMask_32BitDecimalWord)
		  // Run for 1 Single Range
		  Redim Array_RangeInfo(0,3)
		  Array_RangeInfo(0,0) = NetworkSubnetID
		  HostFirst_32BitDecimalWord = NetworkSubnetID+1
		  Array_RangeInfo(0,1) = HostFirst_32BitDecimalWord
		  
		  ReverseSubnetMask = Bitwise.OnesComplement(Input_SubnetMask_32BitDecimalWord)
		  BroadCastID_32BitDecimalWord = NetworkSubnetID+ReverseSubnetMask
		  Array_RangeInfo(0,3) = BroadCastID_32BitDecimalWord
		  
		  HostLast_32BitDecimalWord  = BroadCastID_32BitDecimalWord-1
		  Array_RangeInfo(0,2) =HostLast_32BitDecimalWord
		  
		  
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub fGetSubnetRanges(optional Input_StartIP_32BitDecimalWord as UInt32, optional Input_SubnetMask_32BitDecimalWord as uInt32)
		  #pragma BackgroundTasks false
		  #pragma BoundsChecking false
		  #pragma NilObjectChecking false
		  #pragma StackOverflowChecking false
		  #pragma DisableBackgroundTasks
		  
		  Dim y, i as Integer
		  mCalculateClassFullPrefix
		  
		  Dim NetworkID_Dec, FirstIP_Dec, LastIP_Dec, Bcast_Dec as String
		  
		  if MainWindow.Calc_AllRanges1.User_UseSingleRangeOnly = False Then
		    
		    NetworkSubnetID = fGetClassFullNetwork(Input_StartIP_32BitDecimalWord,Input_SubnetMask_32BitDecimalWord)
		    Redim Array_RangeInfo(Subnets-1,3)
		    Redim Array_DecRangeInfo(Subnets-1,3)
		    CreateArrayThread = New CreateArray_Thread
		    CreateArrayThread.Priority = Thread.HighPriority
		    CreateArrayThread.Run
		    
		    //Preload 40 Lines into Listbox
		    if Subnets <= 200 Then
		      mLoadListbox(Subnets)
		    Else
		      mLoadListbox(300)
		    End if
		    
		    
		  Elseif MainWindow.Calc_AllRanges1.User_UseSingleRangeOnly = True Then
		    NetworkSubnetID = Input_StartIP_32BitDecimalWord AND Input_SubnetMask_32BitDecimalWord
		    mGetClassOnly(Input_StartIP_32BitDecimalWord,Input_SubnetMask_32BitDecimalWord)
		    // Run for 1 Single Range
		    Redim Array_RangeInfo(0,3)
		    Redim Array_DecRangeInfo(0,3)
		    Array_RangeInfo(0,0) = NetworkSubnetID
		    Array_DecRangeInfo(0,0) =  fConvert_32BitDecimalTo8Bit_IP(Array_RangeInfo(0,0))
		    
		    HostFirst_32BitDecimalWord = NetworkSubnetID+1
		    Array_RangeInfo(0,1) = HostFirst_32BitDecimalWord
		    Array_DecRangeInfo(0,1) =fConvert_32BitDecimalTo8Bit_IP(Array_RangeInfo(0,1))
		    ReverseSubnetMask = Bitwise.OnesComplement(Input_SubnetMask_32BitDecimalWord)
		    BroadCastID_32BitDecimalWord = NetworkSubnetID+ReverseSubnetMask
		    Array_RangeInfo(0,3) = BroadCastID_32BitDecimalWord
		    Array_DecRangeInfo(0,3) =fConvert_32BitDecimalTo8Bit_IP(Array_RangeInfo(0,3))
		    
		    HostLast_32BitDecimalWord  = BroadCastID_32BitDecimalWord-1
		    Array_RangeInfo(0,2) =HostLast_32BitDecimalWord
		    Array_DecRangeInfo(0,2) = fConvert_32BitDecimalTo8Bit_IP(Array_RangeInfo(0,2))
		    
		    
		    // Load 1
		    mLoadListbox(1)
		  End if
		  
		  
		  
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub GetSingleRange(NetworkIP as string, SubnetMask as String, Optional ViewMode as String)
		  // Break Down Start IP into Separate Octets in Decimal Form
		  StartIP_1Dec = CDbl(NetworkIP.NthField(".",1))
		  StartIP_2Dec = CDbl(NetworkIP.NthField(".",2))
		  StartIP_3Dec = CDbl(NetworkIP.NthField(".",3))
		  StartIP_4Dec = CDbl(NetworkIP.NthField(".",4))
		  
		  // Break Down Subnet Mask into Separate Octets in Decimal Form
		  SubnetMask1Dec = CDbl(SubnetMask.NthField(".",1))
		  SubnetMask2Dec = CDbl(SubnetMask.NthField(".",2))
		  SubnetMask3Dec = CDbl(SubnetMask.NthField(".",3))
		  SubnetMask4Dec = CDbl(SubnetMask.NthField(".",4))
		  
		  // Convert IP Address and Subnet Mask into 32 Bit Decimal Words for easier processing
		  StartIP_32BitDecimalWord = fConvert_8BitDecimalTo32BitDecimal(StartIP_1Dec, StartIP_2Dec, StartIP_3Dec, StartIP_4Dec)
		  SubnetMask_32BitDecimalWord = fConvert_8BitDecimalTo32BitDecimal(SubnetMask1Dec, SubnetMask2Dec, SubnetMask3Dec, SubnetMask4Dec)
		  
		  
		  //Calulate the Subnet and Host Counts
		  mCalculateNumberOfSubnets
		  mCalculateNumberOfAvailableHosts
		  
		  // Calculate the Actual Subnet Network, Hosts, and Broadcast Addresses and Convert from Binary to Decimal in IPv4 Format
		  // This function will Store Network, All Individual Hosts, and Broadcast address in the 'SubnetCalculator_Storage Class'
		  fGetSingleRanges(StartIP_32BitDecimalWord, SubnetMask_32BitDecimalWord)
		  
		  // Conver the OnesComplement Wildcard mask into IPv4 format
		  mConvertWildcardMask
		  
		  MainWindow.Calc_AllRanges1.SubnetMaskPrefix_Label_Data.Text = "/"+Str(SubnetPrefixLength)
		  MainWindow.Calc_AllRanges1.IANA_Class_Label_Data.Text = IANA_Class
		  
		  Return
		  
		  
		  
		  
		  
		  
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub GetSubnetRanges(NetworkIP as string, SubnetMask as String, Optional ViewMode as String)
		  // Break Down Start IP into Separate Octets in Decimal Form
		  StartIP_1Dec = CDbl(NetworkIP.NthField(".",1))
		  StartIP_2Dec = CDbl(NetworkIP.NthField(".",2))
		  StartIP_3Dec = CDbl(NetworkIP.NthField(".",3))
		  StartIP_4Dec = CDbl(NetworkIP.NthField(".",4))
		  
		  // Break Down Subnet Mask into Separate Octets in Decimal Form
		  SubnetMask1Dec = CDbl(SubnetMask.NthField(".",1))
		  SubnetMask2Dec = CDbl(SubnetMask.NthField(".",2))
		  SubnetMask3Dec = CDbl(SubnetMask.NthField(".",3))
		  SubnetMask4Dec = CDbl(SubnetMask.NthField(".",4))
		  
		  // Convert IP Address and Subnet Mask into 32 Bit Decimal Words for easier processing
		  StartIP_32BitDecimalWord = fConvert_8BitDecimalTo32BitDecimal(StartIP_1Dec, StartIP_2Dec, StartIP_3Dec, StartIP_4Dec)
		  SubnetMask_32BitDecimalWord = fConvert_8BitDecimalTo32BitDecimal(SubnetMask1Dec, SubnetMask2Dec, SubnetMask3Dec, SubnetMask4Dec)
		  
		  
		  //Calulate the Subnet and Host Counts
		  mCalculateNumberOfSubnets
		  mCalculateNumberOfAvailableHosts
		  
		  // Calculate the Actual Subnet Network, Hosts, and Broadcast Addresses and Convert from Binary to Decimal in IPv4 Format
		  // This function will Store Network, All Individual Hosts, and Broadcast address in the 'SubnetCalculator_Storage Class'
		  fGetSubnetRanges(StartIP_32BitDecimalWord, SubnetMask_32BitDecimalWord)
		  
		  // Conver the OnesComplement Wildcard mask into IPv4 format
		  mConvertWildcardMask
		  
		  MainWindow.Calc_AllRanges1.SubnetMaskPrefix_Label_Data.Text = "/"+Str(SubnetPrefixLength)
		  MainWindow.Calc_AllRanges1.IANA_Class_Label_Data.Text = IANA_Class
		  
		  Return
		  
		  
		  
		  
		  
		  
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Sub mCalculateClassFullPrefix()
		  
		  
		  if StartIP_32BitDecimalWord >= 16777216 AND StartIP_32BitDecimalWord <= 2130706431 Then
		    // Class A 1 - 126
		    // Compare our Network Against 255.0.0.0
		    ClassFullSubnetPrefix = 8
		  Else
		    // Class B 128 - 191
		    if StartIP_32BitDecimalWord >= 2147483648 AND StartIP_32BitDecimalWord <= 3221225471 Then
		      // Matched B
		      ClassFullSubnetPrefix = 16
		    else
		      // Class C 192-223
		      If StartIP_32BitDecimalWord >= 3221225472 AND StartIP_32BitDecimalWord <= 3758096383 Then
		        // Matched C
		        ClassFullSubnetPrefix = 24
		      Else
		        // D (Multicast) 224 - 239
		        If StartIP_32BitDecimalWord >= 3758096384 AND StartIP_32BitDecimalWord <= 4026531839 Then
		          // Matched D (Multicast)
		          ClassFullSubnetPrefix = 32
		        End If
		        
		        
		      End if
		      
		    end if
		    
		    
		  end if
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Sub mCalculateNumberOfAvailableHosts()
		  // Calculate the Amount of Total Usable Hosts Per Subnet
		  Dim CalculateHostBits As Integer = fCalculateHostBits
		  HostsPerSubnet = (2 ^CalculateHostBits)-2
		  
		  // Push to Labels on GUI
		  MainWindow.Calc_AllRanges1.TotalNumOfHosts_Data.Text = Format(HostsPerSubnet, "###,###,###")
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Sub mCalculateNumberOfSubnets()
		  //Calulate the Subnet Stuff
		  BitsStolen = fCalculateSubnetBits
		  Subnets = 2 ^ BitsStolen
		  
		  // Push to Labels on GUI
		  MainWindow.Calc_AllRanges1.TotalNumOfSubnets_Data.Text = Format(Subnets,"###,###,###")
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Sub mConvertWildcardMask()
		  // Convert the Cisco Wildcard Inverse Mask to IPv4 format and post it on the GUI label
		  Dim ReverseMask_IPFormat as String
		  ReverseMask_IPFormat = fConvert_32BitDecimalTo8Bit_IP(ReverseSubnetMask)
		  MainWindow.Calc_AllRanges1.CiscoWildcardMask_Data_Label.Text = ReverseMask_IPFormat
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Sub mGetClassOnly(Input_StartIP_32BitDecimalWord as UInt32, Input_SubnetMask_32BitDecimalWord as uInt32)
		  // Get the ClassFull Network Based on the Network ID Passed In - Not Comparing it to the Subnet Mask
		  if Input_StartIP_32BitDecimalWord >= 16777216 AND Input_StartIP_32BitDecimalWord <= 2130706431 Then
		    // Class A 1 - 126
		    // Compare our Network Against 255.0.0.0
		    IANA_Class = "Class A"
		  Else
		    // Class B 128 - 191
		    if Input_StartIP_32BitDecimalWord >= 2147483648 AND Input_StartIP_32BitDecimalWord <= 3221225471 Then
		      // Matched B
		      IANA_Class = "Class B"
		    else
		      // Class C 192-223
		      If Input_StartIP_32BitDecimalWord >= 3221225472 AND Input_StartIP_32BitDecimalWord <= 3758096383 Then
		        // Matched C
		        IANA_Class = "Class C"
		      Else
		        // D (Multicast) 224 - 239
		        If Input_StartIP_32BitDecimalWord >= 3758096384 AND Input_StartIP_32BitDecimalWord <= 4026531839 Then
		          // Matched D (Multicast)
		          IANA_Class = "Class D Multicast"
		        End If
		        
		        
		      End if
		      
		    end if
		    
		    
		  end if
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub mLoadListbox(optional NumberOfLines as Integer)
		  #pragma BackgroundTasks false
		  #pragma BoundsChecking false
		  #pragma NilObjectChecking false
		  #pragma StackOverflowChecking false
		  #pragma DisableBackgroundTasks
		  
		  if MainWindow.Calc_AllRanges1.User_UseSingleRangeOnly = False Then
		    Dim i as Integer
		    EndLoadVar = NumberOfLines + ListLoad_Tracker
		    
		    if EndLoadVar >= Subnets-1 Then
		      MainWindow.LazyLoadScrollBar.Visible = True
		      MainWindow.LazyLoadTimer.Mode = Timer.ModeOff
		    End if
		    
		    // Convert Array Info from 32 Bit Decimal Words into Decimal IPv4 Format
		    for i = ListLoad_Tracker to EndLoadVar-1
		      if i = Subnets - 1 Then
		        Dim Overage as Integer = Abs((EndLoadVar)- Subnets)
		        EndLoadVar = i
		      End if
		      
		      // Update UI
		      dim l as listbox = MainWindow.SubnetRangeListbox
		      l.AddRow(Str(MainWindow.SubnetRangeListbox.LastIndex+2),Array_DecRangeInfo(i,0),Array_DecRangeInfo(i,1),Array_DecRangeInfo(i,2),Array_DecRangeInfo(i,3))
		      dim row as integer = l.lastindex
		      l.CellTag(row,1) =(Array_RangeInfo(i,0))
		      l.CellTag(row,2) =(Array_RangeInfo(i,1))
		      l.CellTag(row,3) =(Array_RangeInfo(i,2))
		      l.CellTag(row,4) =(Array_RangeInfo(i,3))
		    Next i
		    
		    if ListLoad_Tracker >= Subnets Then
		      FinishedLoading = True
		    End if
		    
		    //Increment Global List Load Tracker
		    ListLoad_Tracker = ListLoad_Tracker + NumberOfLines
		    
		  Elseif MainWindow.Calc_AllRanges1.User_UseSingleRangeOnly = True Then
		    // Update UI
		    dim l as listbox = MainWindow.SubnetRangeListbox
		    l.DeleteAllRows
		    l.AddRow("1",Array_DecRangeInfo(0,0),Array_DecRangeInfo(0,1),Array_DecRangeInfo(0,2),Array_DecRangeInfo(0,3))
		    dim row as integer = l.lastindex
		    l.CellTag(row,1) =(Array_RangeInfo(0,0))
		    l.CellTag(row,2) =(Array_RangeInfo(0,1))
		    l.CellTag(row,3) =(Array_RangeInfo(0,2))
		    l.CellTag(row,4) =(Array_RangeInfo(0,3))
		  end if
		  
		  
		End Sub
	#tag EndMethod


	#tag Hook, Flags = &h0
		Event GetAllSubnetRange_Results()
	#tag EndHook

	#tag Hook, Flags = &h0
		Event GetSingleSubnetRange_Results(NetworkID as String, FirstIP as String, LastIP as String, BroadCastIP as String)
	#tag EndHook

	#tag Hook, Flags = &h0
		Event GetTotalNumberOfSubnets(NumberOfSubnets as Integer)
	#tag EndHook

	#tag Hook, Flags = &h0
		Event GetTotalNumberOfUsableHosts(NumberOfUsableHosts as Integer)
	#tag EndHook


	#tag Property, Flags = &h0
		Array_DecRangeInfo(-1,-1) As String
	#tag EndProperty

	#tag Property, Flags = &h0
		Array_RangeInfo(-1,-1) As UInt32
	#tag EndProperty

	#tag Property, Flags = &h0
		BitsStolen As Integer
	#tag EndProperty

	#tag Property, Flags = &h0
		BroadcastID_32BitDecimalWord As uInt32
	#tag EndProperty

	#tag Property, Flags = &h21
		Private ClassFullSubnetPrefix As Integer
	#tag EndProperty

	#tag Property, Flags = &h0
		CreateArrayThread As CreateArray_Thread
	#tag EndProperty

	#tag Property, Flags = &h0
		EndLoadVar As Integer = 300
	#tag EndProperty

	#tag Property, Flags = &h0
		FinishedLoading As Boolean = False
	#tag EndProperty

	#tag Property, Flags = &h21
		Private HostAddressDecimal_First As String
	#tag EndProperty

	#tag Property, Flags = &h21
		Private HostAddressDecimal_Last As String
	#tag EndProperty

	#tag Property, Flags = &h0
		HostFirst_32BitDecimalWord As uInt32
	#tag EndProperty

	#tag Property, Flags = &h0
		HostLast_32BitDecimalWord As uInt32
	#tag EndProperty

	#tag Property, Flags = &h21
		Private HostsPerSubnet As Integer
	#tag EndProperty

	#tag Property, Flags = &h21
		Private IANA_Class As String
	#tag EndProperty

	#tag Property, Flags = &h0
		ListLoad_Tracker As Integer = 0
	#tag EndProperty

	#tag Property, Flags = &h0
		NetworkSubnetID As uInt32
	#tag EndProperty

	#tag Property, Flags = &h0
		ReverseSubnetMask As uInt32
	#tag EndProperty

	#tag Property, Flags = &h0
		SelectData As RecordSet
	#tag EndProperty

	#tag Property, Flags = &h21
		Private StartIP_1Dec As Integer
	#tag EndProperty

	#tag Property, Flags = &h21
		Private StartIP_2Dec As Integer
	#tag EndProperty

	#tag Property, Flags = &h0
		StartIP_32BitDecimalWord As UInt32
	#tag EndProperty

	#tag Property, Flags = &h21
		Private StartIP_3Dec As Integer
	#tag EndProperty

	#tag Property, Flags = &h21
		Private StartIP_4Dec As Integer
	#tag EndProperty

	#tag Property, Flags = &h0
		SubnetEntries_Dictionary_Array() As Dictionary
	#tag EndProperty

	#tag Property, Flags = &h21
		Private SubnetMask1Dec As Integer
	#tag EndProperty

	#tag Property, Flags = &h21
		Private SubnetMask2Dec As Integer
	#tag EndProperty

	#tag Property, Flags = &h21
		Private SubnetMask3Dec As Integer
	#tag EndProperty

	#tag Property, Flags = &h21
		Private SubnetMask4Dec As Integer
	#tag EndProperty

	#tag Property, Flags = &h0
		SubnetMask_32BitDecimalWord As UInt32
	#tag EndProperty

	#tag Property, Flags = &h21
		Private SubnetPrefixLength As Integer
	#tag EndProperty

	#tag Property, Flags = &h0
		Subnets As Integer
	#tag EndProperty


	#tag ViewBehavior
		#tag ViewProperty
			Name="BitsStolen"
			Group="Behavior"
			Type="Integer"
		#tag EndViewProperty
		#tag ViewProperty
			Name="EndLoadVar"
			Group="Behavior"
			InitialValue="300"
			Type="Integer"
		#tag EndViewProperty
		#tag ViewProperty
			Name="FinishedLoading"
			Group="Behavior"
			InitialValue="False"
			Type="Boolean"
		#tag EndViewProperty
		#tag ViewProperty
			Name="Index"
			Visible=true
			Group="ID"
			InitialValue="-2147483648"
			Type="Integer"
		#tag EndViewProperty
		#tag ViewProperty
			Name="Left"
			Visible=true
			Group="Position"
			InitialValue="0"
			Type="Integer"
		#tag EndViewProperty
		#tag ViewProperty
			Name="ListLoad_Tracker"
			Group="Behavior"
			InitialValue="0"
			Type="Integer"
		#tag EndViewProperty
		#tag ViewProperty
			Name="Name"
			Visible=true
			Group="ID"
			Type="String"
		#tag EndViewProperty
		#tag ViewProperty
			Name="Subnets"
			Group="Behavior"
			Type="Integer"
		#tag EndViewProperty
		#tag ViewProperty
			Name="Super"
			Visible=true
			Group="ID"
			Type="String"
		#tag EndViewProperty
		#tag ViewProperty
			Name="Top"
			Visible=true
			Group="Position"
			InitialValue="0"
			Type="Integer"
		#tag EndViewProperty
	#tag EndViewBehavior
End Class
#tag EndClass
