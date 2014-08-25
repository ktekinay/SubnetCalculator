#tag Class
Protected Class App
Inherits Application
	#tag Event
		Sub Open()
		  MainWindow.Show
		End Sub
	#tag EndEvent


	#tag Property, Flags = &h0
		ApplicationStartPositionLeft As Integer = 10
	#tag EndProperty

	#tag Property, Flags = &h0
		ApplicationStartPositionTop As Integer = 60
	#tag EndProperty

	#tag Property, Flags = &h0
		BottomBarTransparencyValue As Integer = 30
	#tag EndProperty

	#tag Property, Flags = &h0
		ThemeColor_BottomBar As Color
	#tag EndProperty

	#tag Property, Flags = &h0
		ThemeColor_Listbox_Color1 As Color = &c7FC07F
	#tag EndProperty

	#tag Property, Flags = &h0
		ThemeColor_Listbox_Color2 As Color = &cDCEEDC
	#tag EndProperty


	#tag Constant, Name = kEditClear, Type = String, Dynamic = False, Default = \"&Delete", Scope = Public
		#Tag Instance, Platform = Windows, Language = Default, Definition  = \"&Delete"
		#Tag Instance, Platform = Linux, Language = Default, Definition  = \"&Delete"
	#tag EndConstant

	#tag Constant, Name = kFileQuit, Type = String, Dynamic = False, Default = \"&Quit", Scope = Public
		#Tag Instance, Platform = Windows, Language = Default, Definition  = \"E&xit"
	#tag EndConstant

	#tag Constant, Name = kFileQuitShortcut, Type = String, Dynamic = False, Default = \"", Scope = Public
		#Tag Instance, Platform = Mac OS, Language = Default, Definition  = \"Cmd+Q"
		#Tag Instance, Platform = Linux, Language = Default, Definition  = \"Ctrl+Q"
	#tag EndConstant


	#tag ViewBehavior
		#tag ViewProperty
			Name="ApplicationStartPositionLeft"
			Group="Behavior"
			InitialValue="10"
			Type="Integer"
		#tag EndViewProperty
		#tag ViewProperty
			Name="ApplicationStartPositionTop"
			Group="Behavior"
			InitialValue="60"
			Type="Integer"
		#tag EndViewProperty
		#tag ViewProperty
			Name="BottomBarTransparencyValue"
			Group="Behavior"
			InitialValue="30"
			Type="Integer"
		#tag EndViewProperty
		#tag ViewProperty
			Name="IP_MyInfo_API_URL"
			Group="Behavior"
			InitialValue="http://ipinfo.io/json"
			Type="string"
			EditorType="MultiLineEditor"
		#tag EndViewProperty
		#tag ViewProperty
			Name="LicenseStatus"
			Group="Behavior"
			Type="String"
			EditorType="MultiLineEditor"
		#tag EndViewProperty
		#tag ViewProperty
			Name="printerSettings"
			Group="Behavior"
			Type="String"
			EditorType="MultiLineEditor"
		#tag EndViewProperty
		#tag ViewProperty
			Name="ThemeColor_BottomBar"
			Group="Behavior"
			InitialValue="&c7FC07F"
			Type="Color"
		#tag EndViewProperty
		#tag ViewProperty
			Name="ThemeColor_Listbox_Color1"
			Group="Behavior"
			InitialValue="&cDCEEDC"
			Type="Color"
		#tag EndViewProperty
		#tag ViewProperty
			Name="ThemeColor_Listbox_Color2"
			Group="Behavior"
			InitialValue="&cFFFFFF"
			Type="Color"
		#tag EndViewProperty
		#tag ViewProperty
			Name="TinyViewEnabled"
			Group="Behavior"
			InitialValue="False"
			Type="Boolean"
		#tag EndViewProperty
		#tag ViewProperty
			Name="UserEnabledOTF_TinyViewEnabled"
			Group="Behavior"
			Type="Boolean"
		#tag EndViewProperty
		#tag ViewProperty
			Name="VersionRelease"
			Group="Behavior"
			Type="String"
			EditorType="MultiLineEditor"
		#tag EndViewProperty
	#tag EndViewBehavior
End Class
#tag EndClass
