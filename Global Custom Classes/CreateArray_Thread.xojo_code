#tag Class
Protected Class CreateArray_Thread
Inherits Thread
	#tag Event
		Sub Run()
		  //#pragma BackgroundTasks false
		  #pragma BoundsChecking false
		  #pragma NilObjectChecking false
		  #pragma StackOverflowChecking false
		  //#pragma DisableBackgroundTasks
		  
		  if MainWindow.Calc_AllRanges1.User_UseSingleRangeOnly = False Then
		    
		    Dim y as Integer
		    Dim NetworkID, FirstIP,LastIP,BroadCastIP as String
		    for y = 0 to MainWindow.SubnetCalculatorSubClass1.Subnets-1
		      
		      // Network Subnet
		      MainWindow.SubnetCalculatorSubClass1.Array_RangeInfo(y,0) = MainWindow.SubnetCalculatorSubClass1.NetworkSubnetID
		      MainWindow.SubnetCalculatorSubClass1.Array_DecRangeInfo(y,0) =  MainWindow.SubnetCalculatorSubClass1.fConvert_32BitDecimalTo8Bit_IP(MainWindow.SubnetCalculatorSubClass1.Array_RangeInfo(y,0))
		      
		      // First Host IP
		      MainWindow.SubnetCalculatorSubClass1.HostFirst_32BitDecimalWord = MainWindow.SubnetCalculatorSubClass1.NetworkSubnetID+1
		      MainWindow.SubnetCalculatorSubClass1.Array_RangeInfo(y,1) =MainWindow.SubnetCalculatorSubClass1.HostFirst_32BitDecimalWord
		      MainWindow.SubnetCalculatorSubClass1.Array_DecRangeInfo(y,1) =MainWindow.SubnetCalculatorSubClass1.fConvert_32BitDecimalTo8Bit_IP(MainWindow.SubnetCalculatorSubClass1.Array_RangeInfo(y,1))
		      
		      // Reverse Mask
		      MainWindow.SubnetCalculatorSubClass1.ReverseSubnetMask = Bitwise.OnesComplement(MainWindow.SubnetCalculatorSubClass1.SubnetMask_32BitDecimalWord)
		      MainWindow.SubnetCalculatorSubClass1.BroadCastID_32BitDecimalWord = MainWindow.SubnetCalculatorSubClass1.NetworkSubnetID+MainWindow.SubnetCalculatorSubClass1.ReverseSubnetMask
		      
		      // BroadCast IP
		      MainWindow.SubnetCalculatorSubClass1.Array_RangeInfo(y,3) = MainWindow.SubnetCalculatorSubClass1.BroadCastID_32BitDecimalWord
		      MainWindow.SubnetCalculatorSubClass1.HostLast_32BitDecimalWord  = MainWindow.SubnetCalculatorSubClass1.BroadCastID_32BitDecimalWord-1
		      MainWindow.SubnetCalculatorSubClass1.Array_DecRangeInfo(y,3) =MainWindow.SubnetCalculatorSubClass1.fConvert_32BitDecimalTo8Bit_IP(MainWindow.SubnetCalculatorSubClass1.Array_RangeInfo(y,3))
		      
		      // Last Host IP
		      MainWindow.SubnetCalculatorSubClass1.Array_RangeInfo(y,2) =MainWindow.SubnetCalculatorSubClass1.HostLast_32BitDecimalWord
		      MainWindow.SubnetCalculatorSubClass1.Array_DecRangeInfo(y,2) = MainWindow.SubnetCalculatorSubClass1.fConvert_32BitDecimalTo8Bit_IP(MainWindow.SubnetCalculatorSubClass1.Array_RangeInfo(y,2))
		      
		      // Increment
		      MainWindow.SubnetCalculatorSubClass1.NetworkSubnetID = MainWindow.SubnetCalculatorSubClass1.BroadCastID_32BitDecimalWord+1
		      
		    Next y
		    
		  end if
		End Sub
	#tag EndEvent


	#tag ViewBehavior
		#tag ViewProperty
			Name="Index"
			Visible=true
			Group="ID"
			Type="Integer"
		#tag EndViewProperty
		#tag ViewProperty
			Name="Left"
			Visible=true
			Group="Position"
			Type="Integer"
		#tag EndViewProperty
		#tag ViewProperty
			Name="Name"
			Visible=true
			Group="ID"
			Type="String"
		#tag EndViewProperty
		#tag ViewProperty
			Name="Priority"
			Visible=true
			Group="Behavior"
			InitialValue="5"
			Type="Integer"
		#tag EndViewProperty
		#tag ViewProperty
			Name="StackSize"
			Visible=true
			Group="Behavior"
			InitialValue="0"
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
			Type="Integer"
		#tag EndViewProperty
	#tag EndViewBehavior
End Class
#tag EndClass
