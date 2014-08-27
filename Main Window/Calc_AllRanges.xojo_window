#tag Window
Begin ContainerControl Calc_AllRanges
   AcceptFocus     =   False
   AcceptTabs      =   True
   AutoDeactivate  =   True
   BackColor       =   &cFFFFFF00
   Backdrop        =   0
   Compatibility   =   ""
   Enabled         =   True
   EraseBackground =   True
   HasBackColor    =   False
   Height          =   148
   HelpTag         =   ""
   InitialParent   =   ""
   Left            =   0
   LockBottom      =   False
   LockLeft        =   True
   LockRight       =   True
   LockTop         =   True
   TabIndex        =   0
   TabPanelIndex   =   0
   TabStop         =   True
   Top             =   0
   Transparent     =   True
   UseFocusRing    =   False
   Visible         =   True
   Width           =   610
   Begin Canvas MiddleBackground_Canvas
      AcceptFocus     =   False
      AcceptTabs      =   False
      AutoDeactivate  =   True
      Backdrop        =   0
      DoubleBuffer    =   True
      Enabled         =   True
      EraseBackground =   True
      Height          =   149
      HelpTag         =   ""
      Index           =   -2147483648
      InitialParent   =   ""
      Left            =   0
      LockBottom      =   False
      LockedInPosition=   False
      LockLeft        =   True
      LockRight       =   False
      LockTop         =   True
      Scope           =   0
      TabIndex        =   0
      TabPanelIndex   =   0
      TabStop         =   True
      Top             =   -1
      Transparent     =   True
      UseFocusRing    =   False
      Visible         =   True
      Width           =   610
      Begin RoundRectangle EnterInfo_BackRetangle
         AutoDeactivate  =   True
         BorderColor     =   &cA8A8A800
         BorderWidth     =   1
         Enabled         =   True
         FillColor       =   &cDDDDDD00
         Height          =   123
         HelpTag         =   ""
         Index           =   -2147483648
         InitialParent   =   "MiddleBackground_Canvas"
         Left            =   10
         LockBottom      =   False
         LockedInPosition=   False
         LockLeft        =   True
         LockRight       =   False
         LockTop         =   True
         OvalHeight      =   16
         OvalWidth       =   10
         Scope           =   0
         TabIndex        =   0
         TabPanelIndex   =   0
         Top             =   12
         Visible         =   True
         Width           =   292
         Begin TextField ValidateAllRange_SubnetMask_TextField
            AcceptTabs      =   False
            Alignment       =   0
            AutoDeactivate  =   True
            AutomaticallyCheckSpelling=   False
            BackColor       =   &cFFFFFF00
            Bold            =   False
            Border          =   True
            CueText         =   " /xx or y.y.y.y"
            DataField       =   ""
            DataSource      =   ""
            Enabled         =   True
            Format          =   ""
            Height          =   22
            HelpTag         =   ""
            Index           =   -2147483648
            InitialParent   =   "EnterInfo_BackRetangle"
            Italic          =   False
            Left            =   133
            LimitText       =   0
            LockBottom      =   False
            LockedInPosition=   False
            LockLeft        =   True
            LockRight       =   False
            LockTop         =   True
            Mask            =   ""
            Password        =   False
            ReadOnly        =   False
            Scope           =   0
            TabIndex        =   1
            TabPanelIndex   =   0
            TabStop         =   True
            Text            =   ""
            TextColor       =   &c00000000
            TextFont        =   "Helvetica"
            TextSize        =   0.0
            TextUnit        =   0
            Top             =   48
            Underline       =   False
            UseFocusRing    =   True
            Visible         =   True
            Width           =   159
         End
         Begin TextField ValidateAllRange_IP_TextField
            AcceptTabs      =   False
            Alignment       =   0
            AutoDeactivate  =   True
            AutomaticallyCheckSpelling=   False
            BackColor       =   &cFFFFFF00
            Bold            =   False
            Border          =   True
            CueText         =   ""
            DataField       =   ""
            DataSource      =   ""
            Enabled         =   True
            Format          =   ""
            Height          =   22
            HelpTag         =   ""
            Index           =   -2147483648
            InitialParent   =   "EnterInfo_BackRetangle"
            Italic          =   False
            Left            =   133
            LimitText       =   0
            LockBottom      =   False
            LockedInPosition=   False
            LockLeft        =   True
            LockRight       =   False
            LockTop         =   True
            Mask            =   ""
            Password        =   False
            ReadOnly        =   False
            Scope           =   0
            TabIndex        =   0
            TabPanelIndex   =   0
            TabStop         =   True
            Text            =   ""
            TextColor       =   &c00000000
            TextFont        =   "Helvetica"
            TextSize        =   0.0
            TextUnit        =   0
            Top             =   22
            Underline       =   False
            UseFocusRing    =   True
            Visible         =   True
            Width           =   159
         End
         Begin Label EnterIPAddressLabel
            AutoDeactivate  =   True
            Bold            =   False
            DataField       =   ""
            DataSource      =   ""
            Enabled         =   True
            Height          =   20
            HelpTag         =   ""
            Index           =   -2147483648
            InitialParent   =   "EnterInfo_BackRetangle"
            Italic          =   False
            Left            =   21
            LockBottom      =   False
            LockedInPosition=   False
            LockLeft        =   True
            LockRight       =   False
            LockTop         =   True
            Multiline       =   False
            Scope           =   0
            Selectable      =   False
            TabIndex        =   2
            TabPanelIndex   =   0
            Text            =   "#kMainCalcAllRangesIP"
            TextAlign       =   0
            TextColor       =   &c00000000
            TextFont        =   "Helvetica"
            TextSize        =   0.0
            TextUnit        =   0
            Top             =   23
            Transparent     =   True
            Underline       =   False
            Visible         =   True
            Width           =   100
         End
         Begin Label ValidateALLRange_SubnetMask_Label
            AutoDeactivate  =   True
            Bold            =   False
            DataField       =   ""
            DataSource      =   ""
            Enabled         =   True
            Height          =   21
            HelpTag         =   ""
            Index           =   -2147483648
            InitialParent   =   "EnterInfo_BackRetangle"
            Italic          =   False
            Left            =   21
            LockBottom      =   False
            LockedInPosition=   False
            LockLeft        =   True
            LockRight       =   False
            LockTop         =   True
            Multiline       =   False
            Scope           =   0
            Selectable      =   False
            TabIndex        =   3
            TabPanelIndex   =   0
            Text            =   "#kMainCalcAllRangesSubnetMask"
            TextAlign       =   0
            TextColor       =   &c00000000
            TextFont        =   "Helvetica"
            TextSize        =   0.0
            TextUnit        =   0
            Top             =   47
            Transparent     =   True
            Underline       =   False
            Visible         =   True
            Width           =   136
         End
         Begin CheckBox SingleRangeCheckbox
            AutoDeactivate  =   True
            Bold            =   False
            Caption         =   "#kMainCalcAllRanges_ShowSpecRangeText"
            DataField       =   ""
            DataSource      =   ""
            Enabled         =   True
            Height          =   20
            HelpTag         =   "This will only show the range that the entered IP Address is a member of"
            Index           =   -2147483648
            InitialParent   =   "EnterInfo_BackRetangle"
            Italic          =   False
            Left            =   21
            LockBottom      =   False
            LockedInPosition=   False
            LockLeft        =   True
            LockRight       =   False
            LockTop         =   True
            Scope           =   0
            State           =   0
            TabIndex        =   2
            TabPanelIndex   =   0
            TabStop         =   True
            TextFont        =   "Helvetica"
            TextSize        =   0.0
            TextUnit        =   0
            Top             =   77
            Underline       =   False
            Value           =   False
            Visible         =   True
            Width           =   172
         End
         Begin PushButton ValidateAllRange_CalculateButton
            AutoDeactivate  =   True
            Bold            =   False
            ButtonStyle     =   "0"
            Cancel          =   False
            Caption         =   "#kMainCalcAllRanges_CalculateButton"
            Default         =   False
            Enabled         =   True
            Height          =   22
            HelpTag         =   ""
            Index           =   -2147483648
            InitialParent   =   "EnterInfo_BackRetangle"
            Italic          =   False
            Left            =   138
            LockBottom      =   False
            LockedInPosition=   False
            LockLeft        =   True
            LockRight       =   False
            LockTop         =   True
            Scope           =   0
            TabIndex        =   4
            TabPanelIndex   =   0
            TabStop         =   True
            TextFont        =   "Helvetica"
            TextSize        =   12.0
            TextUnit        =   0
            Top             =   104
            Underline       =   False
            Visible         =   True
            Width           =   81
         End
      End
      Begin RoundRectangle Results_BackRectangle
         AutoDeactivate  =   True
         BorderColor     =   &cA8A8A800
         BorderWidth     =   1
         Enabled         =   True
         FillColor       =   &cDDDDDD00
         Height          =   123
         HelpTag         =   ""
         Index           =   -2147483648
         InitialParent   =   "MiddleBackground_Canvas"
         Left            =   311
         LockBottom      =   False
         LockedInPosition=   False
         LockLeft        =   True
         LockRight       =   False
         LockTop         =   True
         OvalHeight      =   16
         OvalWidth       =   10
         Scope           =   0
         TabIndex        =   1
         TabPanelIndex   =   0
         Top             =   12
         Visible         =   True
         Width           =   289
         Begin Label NumberofSubnetsTotal_Label
            AutoDeactivate  =   True
            Bold            =   False
            DataField       =   ""
            DataSource      =   ""
            Enabled         =   True
            Height          =   20
            HelpTag         =   ""
            Index           =   -2147483648
            InitialParent   =   "Results_BackRectangle"
            Italic          =   False
            Left            =   320
            LockBottom      =   False
            LockedInPosition=   False
            LockLeft        =   True
            LockRight       =   False
            LockTop         =   True
            Multiline       =   False
            Scope           =   0
            Selectable      =   False
            TabIndex        =   1
            TabPanelIndex   =   0
            Text            =   "#kMainCalcAllRanges_TotalNumberofSubnets"
            TextAlign       =   0
            TextColor       =   &c00000000
            TextFont        =   "Helvetica"
            TextSize        =   0.0
            TextUnit        =   0
            Top             =   18
            Transparent     =   True
            Underline       =   False
            Visible         =   True
            Width           =   165
         End
         Begin Label NumberofHostsTotal_Label
            AutoDeactivate  =   True
            Bold            =   False
            DataField       =   ""
            DataSource      =   ""
            Enabled         =   True
            Height          =   21
            HelpTag         =   ""
            Index           =   -2147483648
            InitialParent   =   "Results_BackRectangle"
            Italic          =   False
            Left            =   320
            LockBottom      =   False
            LockedInPosition=   False
            LockLeft        =   True
            LockRight       =   False
            LockTop         =   True
            Multiline       =   False
            Scope           =   0
            Selectable      =   False
            TabIndex        =   3
            TabPanelIndex   =   0
            Text            =   "#kMainCalcAllRanges_TotalNumberofUsableHosts"
            TextAlign       =   0
            TextColor       =   &c00000000
            TextFont        =   "Helvetica"
            TextSize        =   0.0
            TextUnit        =   0
            Top             =   39
            Transparent     =   True
            Underline       =   False
            Visible         =   True
            Width           =   165
         End
         Begin Label IANAClass_Label
            AutoDeactivate  =   True
            Bold            =   False
            DataField       =   ""
            DataSource      =   ""
            Enabled         =   True
            Height          =   20
            HelpTag         =   ""
            Index           =   -2147483648
            InitialParent   =   "Results_BackRectangle"
            Italic          =   False
            Left            =   321
            LockBottom      =   False
            LockedInPosition=   False
            LockLeft        =   True
            LockRight       =   False
            LockTop         =   True
            Multiline       =   False
            Scope           =   0
            Selectable      =   False
            TabIndex        =   4
            TabPanelIndex   =   0
            Text            =   "#kMainCalcAllRangesIANAClassfulClass"
            TextAlign       =   0
            TextColor       =   &c00000000
            TextFont        =   "Helvetica"
            TextSize        =   0.0
            TextUnit        =   0
            Top             =   63
            Transparent     =   True
            Underline       =   False
            Visible         =   True
            Width           =   165
         End
         Begin Label SubnetPrefixLength
            AutoDeactivate  =   True
            Bold            =   False
            DataField       =   ""
            DataSource      =   ""
            Enabled         =   True
            Height          =   20
            HelpTag         =   ""
            Index           =   -2147483648
            InitialParent   =   "Results_BackRectangle"
            Italic          =   False
            Left            =   320
            LockBottom      =   False
            LockedInPosition=   False
            LockLeft        =   True
            LockRight       =   False
            LockTop         =   True
            Multiline       =   False
            Scope           =   0
            Selectable      =   False
            TabIndex        =   5
            TabPanelIndex   =   0
            Text            =   "#kMailCalcAllRanges_SubnetMaskPrefixLength"
            TextAlign       =   0
            TextColor       =   &c00000000
            TextFont        =   "Helvetica"
            TextSize        =   0.0
            TextUnit        =   0
            Top             =   87
            Transparent     =   True
            Underline       =   False
            Visible         =   True
            Width           =   165
         End
         Begin Label CiscoInverseMask_Label
            AutoDeactivate  =   True
            Bold            =   False
            DataField       =   ""
            DataSource      =   ""
            Enabled         =   True
            Height          =   20
            HelpTag         =   ""
            Index           =   -2147483648
            InitialParent   =   "Results_BackRectangle"
            Italic          =   False
            Left            =   320
            LockBottom      =   False
            LockedInPosition=   False
            LockLeft        =   True
            LockRight       =   False
            LockTop         =   True
            Multiline       =   False
            Scope           =   0
            Selectable      =   False
            TabIndex        =   7
            TabPanelIndex   =   0
            Text            =   ""
            TextAlign       =   0
            TextColor       =   &c00000000
            TextFont        =   "Helvetica"
            TextSize        =   0.0
            TextUnit        =   0
            Top             =   111
            Transparent     =   True
            Underline       =   False
            Visible         =   True
            Width           =   165
         End
         Begin TextField TotalNumOfSubnets_Data
            AcceptTabs      =   False
            Alignment       =   1
            AutoDeactivate  =   True
            AutomaticallyCheckSpelling=   False
            BackColor       =   &cFFFFFF00
            Bold            =   False
            Border          =   True
            CueText         =   ""
            DataField       =   ""
            DataSource      =   ""
            Enabled         =   True
            Format          =   ""
            Height          =   22
            HelpTag         =   ""
            Index           =   -2147483648
            InitialParent   =   "Results_BackRectangle"
            Italic          =   False
            Left            =   479
            LimitText       =   0
            LockBottom      =   False
            LockedInPosition=   False
            LockLeft        =   True
            LockRight       =   True
            LockTop         =   True
            Mask            =   ""
            Password        =   False
            ReadOnly        =   False
            Scope           =   0
            TabIndex        =   10
            TabPanelIndex   =   0
            TabStop         =   True
            Text            =   ""
            TextColor       =   &c1E731100
            TextFont        =   "Helvetica"
            TextSize        =   12.0
            TextUnit        =   0
            Top             =   17
            Underline       =   False
            UseFocusRing    =   True
            Visible         =   True
            Width           =   113
         End
         Begin TextField TotalNumOfHosts_Data
            AcceptTabs      =   False
            Alignment       =   1
            AutoDeactivate  =   True
            AutomaticallyCheckSpelling=   False
            BackColor       =   &cFFFFFF00
            Bold            =   False
            Border          =   True
            CueText         =   ""
            DataField       =   ""
            DataSource      =   ""
            Enabled         =   True
            Format          =   ""
            Height          =   22
            HelpTag         =   ""
            Index           =   -2147483648
            InitialParent   =   "Results_BackRectangle"
            Italic          =   False
            Left            =   479
            LimitText       =   0
            LockBottom      =   False
            LockedInPosition=   False
            LockLeft        =   True
            LockRight       =   True
            LockTop         =   True
            Mask            =   ""
            Password        =   False
            ReadOnly        =   False
            Scope           =   0
            TabIndex        =   11
            TabPanelIndex   =   0
            TabStop         =   True
            Text            =   ""
            TextColor       =   &c1E731100
            TextFont        =   "Helvetica"
            TextSize        =   12.0
            TextUnit        =   0
            Top             =   40
            Underline       =   False
            UseFocusRing    =   True
            Visible         =   True
            Width           =   113
         End
         Begin TextField IANA_Class_Label_Data
            AcceptTabs      =   False
            Alignment       =   1
            AutoDeactivate  =   True
            AutomaticallyCheckSpelling=   False
            BackColor       =   &cFFFFFF00
            Bold            =   False
            Border          =   True
            CueText         =   ""
            DataField       =   ""
            DataSource      =   ""
            Enabled         =   True
            Format          =   ""
            Height          =   22
            HelpTag         =   ""
            Index           =   -2147483648
            InitialParent   =   "Results_BackRectangle"
            Italic          =   False
            Left            =   479
            LimitText       =   0
            LockBottom      =   False
            LockedInPosition=   False
            LockLeft        =   True
            LockRight       =   True
            LockTop         =   True
            Mask            =   ""
            Password        =   False
            ReadOnly        =   False
            Scope           =   0
            TabIndex        =   12
            TabPanelIndex   =   0
            TabStop         =   True
            Text            =   ""
            TextColor       =   &c0000FE00
            TextFont        =   "Helvetica"
            TextSize        =   12.0
            TextUnit        =   0
            Top             =   63
            Underline       =   False
            UseFocusRing    =   True
            Visible         =   True
            Width           =   113
         End
         Begin TextField SubnetMaskPrefix_Label_Data
            AcceptTabs      =   False
            Alignment       =   1
            AutoDeactivate  =   True
            AutomaticallyCheckSpelling=   False
            BackColor       =   &cFFFFFF00
            Bold            =   False
            Border          =   True
            CueText         =   ""
            DataField       =   ""
            DataSource      =   ""
            Enabled         =   True
            Format          =   ""
            Height          =   22
            HelpTag         =   ""
            Index           =   -2147483648
            InitialParent   =   "Results_BackRectangle"
            Italic          =   False
            Left            =   479
            LimitText       =   0
            LockBottom      =   False
            LockedInPosition=   False
            LockLeft        =   True
            LockRight       =   True
            LockTop         =   True
            Mask            =   ""
            Password        =   False
            ReadOnly        =   False
            Scope           =   0
            TabIndex        =   13
            TabPanelIndex   =   0
            TabStop         =   True
            Text            =   ""
            TextColor       =   &c0000FE00
            TextFont        =   "Helvetica"
            TextSize        =   12.0
            TextUnit        =   0
            Top             =   86
            Underline       =   False
            UseFocusRing    =   True
            Visible         =   True
            Width           =   113
         End
         Begin TextField CiscoWildcardMask_Data_Label
            AcceptTabs      =   False
            Alignment       =   1
            AutoDeactivate  =   True
            AutomaticallyCheckSpelling=   False
            BackColor       =   &cFFFFFF00
            Bold            =   False
            Border          =   True
            CueText         =   ""
            DataField       =   ""
            DataSource      =   ""
            Enabled         =   True
            Format          =   ""
            Height          =   22
            HelpTag         =   ""
            Index           =   -2147483648
            InitialParent   =   "Results_BackRectangle"
            Italic          =   False
            Left            =   479
            LimitText       =   0
            LockBottom      =   False
            LockedInPosition=   False
            LockLeft        =   True
            LockRight       =   True
            LockTop         =   True
            Mask            =   ""
            Password        =   False
            ReadOnly        =   False
            Scope           =   0
            TabIndex        =   14
            TabPanelIndex   =   0
            TabStop         =   True
            Text            =   ""
            TextColor       =   &c0000FE00
            TextFont        =   "Helvetica"
            TextSize        =   12.0
            TextUnit        =   0
            Top             =   109
            Underline       =   False
            UseFocusRing    =   True
            Visible         =   True
            Width           =   113
         End
      End
   End
End
#tag EndWindow

#tag WindowCode
	#tag Event
		Sub MouseMove(X As Integer, Y As Integer)
		  ToolTip.Hide
		End Sub
	#tag EndEvent

	#tag Event
		Sub Open()
		  ValidateAllRange_IP_TextField.SetFocus
		End Sub
	#tag EndEvent


	#tag Method, Flags = &h21
		Private Sub mSubmitValidation()
		  // Force one Last IP Address Validation
		  Dim IPV4_Validation as String =MainWindow. fIPv4_Validation(ValidateAllRange_IP_TextField.Text)
		  Select Case IPV4_Validation
		  Case "Not_Valid"
		    Tooltip.Show("Please Enter a Valid IPv4 Address", System.MouseX, System.MouseY, True)
		    ValidateAllRange_IP_TextField.Text = ""
		    Exit
		  End Select
		  
		  // Force one last Subnet Mask Validation
		  Dim SubnetMask_Validation as String = MainWindow.fSubnetMaskValidation(ValidateAllRange_SubnetMask_TextField.Text)
		  Select Case SubnetMask_Validation
		  Case "Not_Valid"
		    Tooltip.Show("Please Enter a Valid Subnet Mask", System.MouseX, System.MouseY, True)
		    ValidateAllRange_SubnetMask_TextField.Text = ""
		    Exit
		  Else
		    ValidateAllRange_SubnetMask_TextField.Text = SubnetMask_Validation
		  End Select
		  
		  // Clear Out Listbox First
		  MainWindow.SubnetRangeListbox.DeleteAllRows
		  // Send the SubnetCalculator GetSubnetRangeWithoutHosts Method
		  MainWindow.SubnetCalculatorSubClass1.GetSubnetRanges(ValidateAllRange_IP_TextField.Text,ValidateAllRange_SubnetMask_TextField.Text)
		  // Begin back at Row 0
		  MainWindow.SubnetRangeListbox.ScrollPosition = 0
		  // Reset Display Stats
		  if MainWindow.SubnetRangeListbox.ListCount >= 11 Then
		    MainWindow.CurrentDisplayCounts_DataLabel.Text = Format(MainWindow.SubnetRangeListbox.ScrollPosition+11,"###,###,###") + " of " + Format(MainWindow.SubnetCalculatorSubClass1.Subnets,"###,###,###")
		  Elseif MainWindow.SubnetRangeListbox.ListCount < 11 Then
		    MainWindow.CurrentDisplayCounts_DataLabel.Text = Format(MainWindow.SubnetCalculatorSubClass1.Subnets,"###,###,###") + " of " + Format(MainWindow.SubnetCalculatorSubClass1.Subnets,"###,###,###")
		  End If
		  
		End Sub
	#tag EndMethod


	#tag Property, Flags = &h0
		BeginOnClassLessSubnet As Boolean = False
	#tag EndProperty

	#tag Property, Flags = &h0
		User_UseSingleRangeOnly As Boolean = False
	#tag EndProperty


#tag EndWindowCode

#tag Events MiddleBackground_Canvas
	#tag Event
		Sub Paint(g As Graphics, areas() As REALbasic.Rect)
		  //g.ForeColor = &cffffff
		  //g.FillRect(0,0,me.Width,me.Height)
		  
		  // Awesome Gradient Fill
		  dim i as integer, ratio, endratio as Double
		  dim StartColor, EndColor as Color
		  g.Transparency = App.BottomBarTransparencyValue
		  
		  StartColor = &cFFFFFF//RGB(235, 239, 242)
		  EndColor = App. ThemeColor_BottomBar//71B971
		  
		  // Draw The Gradient
		  for i = g.Height DownTo 0
		    // Need our ratios of start / end
		    ratio = (i/g.Height)
		    endratio = ((g.Height-i)/g.Height)
		    // Determine the Color
		    g.ForeColor = RGB(StartColor.Red * endratio + EndColor.Red * ratio, StartColor.Green * endratio + EndColor.Green * ratio, StartColor.Blue * endratio + EndColor.Blue * ratio)
		    // Draw the current line
		    g.DrawLine 0, i, g.Width, i
		  next
		End Sub
	#tag EndEvent
#tag EndEvents
#tag Events ValidateAllRange_SubnetMask_TextField
	#tag Event
		Function KeyDown(Key As String) As Boolean
		  Tooltip.Hide
		  If key = Chr(9) Or key = Chr(13) Then 'tab key,right arrow, return 'move to next available entry cell
		    // Run Subnet Mask validation
		    Dim SubnetMask_Validation as String = MainWindow.fSubnetMaskValidation(me.Text)
		    Select Case SubnetMask_Validation
		    Case "Not_Valid"
		      Tooltip.Show("Please Enter a Valid Subnet Mask", System.MouseX, System.MouseY, True)
		      me.Text = ""
		      Return True
		    Else
		      me.Text = SubnetMask_Validation
		    End Select
		  End If
		  
		  If Key = Chr(13) Then
		    ValidateAllRange_CalculateButton.SetFocus
		    if MainWindow.SubnetCalculatorSubClass1 <> Nil Then
		      // Instantiate Subnet Calculator Class
		      MainWindow.SubnetCalculatorSubClass1 = Nil
		    end if
		    
		    MainWindow.SubnetCalculatorSubClass1 = New SubnetCalculator_Class
		    // Perform actions in case this is a Second+ Run
		    MainWindow.LazyLoadScrollBar.Visible =False
		    MainWindow.LazyLoadTimer.Mode = Timer.ModeOff
		    MainWindow.SubnetCalculatorSubClass1.KillCreateArrayThread()
		    
		    mSubmitValidation
		    // Set Lazy Load Scroll Bar Options
		    MainWindow.LazyLoadScrollBar.Maximum = MainWindow.SubnetCalculatorSubClass1.Subnets-1
		    MainWindow.LazyLoadScrollBar.Minimum = 0
		    MainWindow.LazyLoadScrollBar.Value = 0
		    
		    IF MainWindow.Calc_AllRanges1.User_UseSingleRangeOnly = false Then
		      MainWindow.CurrentDisplayCounts_DataLabel.Visible = True
		      MainWindow.Height =409
		      MainWindow.mUpdateDisplayRecordCounter
		      MainWindow.LazyLoadTimer.Period = 900
		      MainWindow.LazyLoadTimer.Mode = Timer.ModeMultiple
		    Elseif MainWindow.Calc_AllRanges1.User_UseSingleRangeOnly = True Then
		      MainWindow.CurrentDisplayCounts_DataLabel.Visible = False
		      MainWindow.Height =265
		    END IF
		    
		    
		    
		    
		    
		    Return true
		  End if
		End Function
	#tag EndEvent
	#tag Event
		Sub MouseMove(X As Integer, Y As Integer)
		  Tooltip.Hide
		End Sub
	#tag EndEvent
	#tag Event
		Sub MouseUp(X As Integer, Y As Integer)
		  // Run Subnet Mask validation
		  Dim SubnetMask_Validation as String = MainWindow.fSubnetMaskValidation(me.Text)
		  Select Case SubnetMask_Validation
		  Case "Not_Valid"
		    Tooltip.Show("Please Enter a Valid Subnet Mask", System.MouseX, System.MouseY, True)
		    me.Text = ""
		  Else
		    me.Text = SubnetMask_Validation
		  End Select
		End Sub
	#tag EndEvent
	#tag Event
		Sub Open()
		  me.AcceptTextDrop
		  
		  #IF TargetMacOS Then
		    Declare Sub setBezelStyle Lib "Cocoa" Selector "setBezelStyle:" (handle As Integer, value As Integer)
		    setBezelStyle(Me.Handle, 1)
		  #Endif
		End Sub
	#tag EndEvent
	#tag Event
		Sub DropObject(obj As DragItem, action As Integer)
		  If Obj.TextAvailable then
		    Me.Text = Obj.Text
		  End if
		End Sub
	#tag EndEvent
#tag EndEvents
#tag Events ValidateAllRange_IP_TextField
	#tag Event
		Function KeyDown(Key As String) As Boolean
		  Tooltip.Hide
		  Select Case key
		  Case Chr(32)
		    Tooltip.Show("Sorry, Spaces not allowed", System.MouseX, System.MouseY, True)
		    Return True
		    
		  Case Chr(33)
		    Tooltip.Show("Sorry, Exclamation mark not allowed", System.MouseX, System.MouseY, True)
		    Return True
		    
		  Case Chr(34)
		    Tooltip.Show("Sorry, Quotation marks not allowed", System.MouseX, System.MouseY, True)
		    Return True
		    
		  Case Chr(35)
		    Tooltip.Show("Sorry, Pound sign not allowed", System.MouseX, System.MouseY, True)
		    Return True
		    
		  Case Chr(36)
		    Tooltip.Show("Sorry, Dollar sign not allowed", System.MouseX, System.MouseY, True)
		    Return True
		    
		  Case Chr(37)
		    Tooltip.Show("Sorry, Percentage sign not allowed", System.MouseX, System.MouseY, True)
		    Return True
		    
		  Case Chr(38)
		    Tooltip.Show("Sorry, Ampersand not allowed", System.MouseX, System.MouseY, True)
		    Return True
		    
		  Case Chr(39)
		    Tooltip.Show("Sorry, Apostrophe not allowed", System.MouseX, System.MouseY, True)
		    Return True
		    
		  Case Chr(40)
		    Tooltip.Show("Sorry, Open Parenthesis not allowed", System.MouseX, System.MouseY, True)
		    Return True
		    
		  Case Chr(41)
		    Tooltip.Show("Sorry, Closed Parenthesis not allowed", System.MouseX, System.MouseY, True)
		    Return True
		    
		  Case Chr(42)
		    Tooltip.Show("Sorry, Asterisk not allowed", System.MouseX, System.MouseY, True)
		    Return True
		    
		  Case Chr(43)
		    Tooltip.Show("Sorry, Plus signs not allowed", System.MouseX, System.MouseY, True)
		    Return True
		    
		  Case Chr(44)
		    Tooltip.Show("Sorry, Commas not allowed", System.MouseX, System.MouseY, True)
		    Return True
		    
		  Case Chr(45)
		    Tooltip.Show("Sorry, Hyphen not allowed", System.MouseX, System.MouseY, True)
		    Return True
		    
		    // Allow 46 (Period)
		    
		  Case Chr(47)
		    Tooltip.Show("Sorry, Only Numbers and Periods allowed", System.MouseX, System.MouseY, True)
		    Return True
		    
		    // Allow 48 - 57 (Numbers)
		    
		  Case Chr(58)
		    Tooltip.Show("Sorry, Colon not allowed", System.MouseX, System.MouseY, True)
		    Return True
		    
		  Case Chr(59)
		    Tooltip.Show("Sorry, Semi Colon not allowed", System.MouseX, System.MouseY, True)
		    Return True
		    
		  Case Chr(60)
		    Tooltip.Show("Sorry, Less than bracket not allowed", System.MouseX, System.MouseY, True)
		    Return True
		    
		  Case Chr(61)
		    Tooltip.Show("Sorry, Equal sign not allowed", System.MouseX, System.MouseY, True)
		    Return True
		    
		  Case Chr(62)
		    Tooltip.Show("Sorry, Greater than bracket not allowed", System.MouseX, System.MouseY, True)
		    Return True
		    
		  Case Chr(63)
		    Tooltip.Show("Sorry, Question Mark not allowed", System.MouseX, System.MouseY, True)
		    Return True
		    
		  Case Chr(64)
		    Tooltip.Show("Sorry, ""At"" signs not allowed", System.MouseX, System.MouseY, True)
		    Return True
		    
		  Case Chr(65) to Chr(90)
		    Tooltip.Show("Sorry, Only Numbers and Periods allowed", System.MouseX, System.MouseY, True)
		    Return True
		    
		  Case Chr(91)
		    Tooltip.Show("Sorry, Open Square Bracket not allowed", System.MouseX, System.MouseY, True)
		    Return True
		    
		  Case Chr(92)
		    Tooltip.Show("Sorry, Back Slash not allowed", System.MouseX, System.MouseY, True)
		    Return True
		    
		  Case Chr(93)
		    Tooltip.Show("Sorry, Closed Square Bracket not allowed", System.MouseX, System.MouseY, True)
		    Return True
		    
		  Case Chr(94)
		    Tooltip.Show("Sorry, Caret not allowed", System.MouseX, System.MouseY, True)
		    Return True
		    
		  Case Chr(95)
		    Tooltip.Show("Sorry, Underscore not allowed", System.MouseX, System.MouseY, True)
		    Return True
		    
		  Case Chr(96)
		    Tooltip.Show("Sorry, Tilde not allowed", System.MouseX, System.MouseY, True)
		    Return True
		    
		  Case Chr(97) to Chr(122)
		    Tooltip.Show("Sorry, Only Numbers and Periods allowed", System.MouseX, System.MouseY, True)
		    Return True
		    
		  Case Chr(123)
		    Tooltip.Show("Sorry, Open Curly Brackets not allowed", System.MouseX, System.MouseY, True)
		    Return True
		    
		  Case Chr(124)
		    Tooltip.Show("Sorry, Pipe not allowed", System.MouseX, System.MouseY, True)
		    Return True
		    
		  Case Chr(125)
		    Tooltip.Show("Sorry, Closed Curly Brackets not allowed", System.MouseX, System.MouseY, True)
		    Return True
		    
		  Case Chr(126)
		    Tooltip.Show("Sorry, Grave Accent Marks not allowed", System.MouseX, System.MouseY, True)
		    Return True
		    
		  Case Chr(127)
		    Tooltip.Show("Sorry, Only Numbers and Periods allowed", System.MouseX, System.MouseY, True)
		    Return True
		  End Select
		  
		  If key = Chr(9) Or key = Chr(13) Then
		    Dim IPV4_Validation as String = MainWindow.fIPv4_Validation(me.Text)
		    Select Case IPV4_Validation
		    Case "Not_Valid"
		      Tooltip.Show("Please Enter a Valid IPv4 Address", System.MouseX, System.MouseY, True)
		      Me.Text = ""
		      Return True
		    End Select
		  End if
		  
		  If Key = Chr(13) Then
		    ValidateAllRange_SubnetMask_TextField.SetFocus
		    
		  End if
		  
		End Function
	#tag EndEvent
	#tag Event
		Sub MouseMove(X As Integer, Y As Integer)
		  ToolTip.Hide
		  
		End Sub
	#tag EndEvent
	#tag Event
		Sub MouseUp(X As Integer, Y As Integer)
		  Dim IPV4_Validation as String = MainWindow.fIPv4_Validation(me.Text)
		  Select Case IPV4_Validation
		  Case "Not_Valid"
		    Tooltip.Show("Please Enter a Valid IPv4 Address", System.MouseX, System.MouseY, True)
		    Me.Text = ""
		  End Select
		End Sub
	#tag EndEvent
	#tag Event
		Sub Open()
		  me.AcceptTextDrop
		  
		  #IF TargetMacOS Then
		    Declare Sub setBezelStyle Lib "Cocoa" Selector "setBezelStyle:" (handle As Integer, value As Integer)
		    setBezelStyle(Me.Handle, 1)
		  #Endif
		End Sub
	#tag EndEvent
	#tag Event
		Sub DropObject(obj As DragItem, action As Integer)
		  If Obj.TextAvailable then
		    Me.Text = Obj.Text
		  End if
		End Sub
	#tag EndEvent
#tag EndEvents
#tag Events SingleRangeCheckbox
	#tag Event
		Sub Action()
		  if me.Value = true then
		    // Is Checked
		    MainWindow.Calc_AllRanges1.User_UseSingleRangeOnly = True
		    MainWindow.LazyLoadScrollBar.Visible = False
		    
		  elseif me.Value = False then
		    MainWindow.Calc_AllRanges1.User_UseSingleRangeOnly = False
		  end if
		End Sub
	#tag EndEvent
	#tag Event
		Sub Open()
		  Me.Value = True
		End Sub
	#tag EndEvent
#tag EndEvents
#tag Events ValidateAllRange_CalculateButton
	#tag Event
		Sub Action()
		  if MainWindow.SubnetCalculatorSubClass1 <> Nil Then
		    // Instantiate Subnet Calculator Class
		    MainWindow.SubnetCalculatorSubClass1 = Nil
		  end if
		  
		  MainWindow.SubnetCalculatorSubClass1 = New SubnetCalculator_Class
		  // Perform actions in case this is a Second+ Run
		  MainWindow.LazyLoadScrollBar.Visible =False
		  MainWindow.LazyLoadTimer.Mode = Timer.ModeOff
		  MainWindow.SubnetCalculatorSubClass1.KillCreateArrayThread()
		  
		  mSubmitValidation
		  // Set Lazy Load Scroll Bar Options
		  MainWindow.LazyLoadScrollBar.Maximum = MainWindow.SubnetCalculatorSubClass1.Subnets-1
		  MainWindow.LazyLoadScrollBar.Minimum = 0
		  MainWindow.LazyLoadScrollBar.Value = 0
		  
		  IF MainWindow.Calc_AllRanges1.User_UseSingleRangeOnly = false Then
		    MainWindow.CurrentDisplayCounts_DataLabel.Visible = True
		    MainWindow.Height =409
		    MainWindow.mUpdateDisplayRecordCounter
		    MainWindow.LazyLoadTimer.Period = 900
		    MainWindow.LazyLoadTimer.Mode = Timer.ModeMultiple
		  Elseif MainWindow.Calc_AllRanges1.User_UseSingleRangeOnly = True Then
		    MainWindow.CurrentDisplayCounts_DataLabel.Visible = False
		    MainWindow.Height =265
		  END IF
		  
		  
		  
		  
		  
		  
		  
		  
		  
		  
		End Sub
	#tag EndEvent
	#tag Event
		Function KeyDown(Key As String) As Boolean
		  If Key = Chr(13) Then
		    
		    
		    
		    
		    
		    Return True
		  end if
		End Function
	#tag EndEvent
#tag EndEvents
#tag Events CiscoInverseMask_Label
	#tag Event
		Sub Open()
		  Me.Text = "Cisco" + Encodings.UTF8.Chr(174) + " Wildcard Mask:"
		End Sub
	#tag EndEvent
#tag EndEvents
#tag Events TotalNumOfSubnets_Data
	#tag Event
		Sub Open()
		  #IF TargetMacOS Then
		    Declare Sub setBezelStyle Lib "Cocoa" Selector "setBezelStyle:" (handle As Integer, value As Integer)
		    setBezelStyle(Me.Handle, 1)
		  #Endif
		  
		End Sub
	#tag EndEvent
	#tag Event
		Function MouseDown(X As Integer, Y As Integer) As Boolean
		  Return True
		End Function
	#tag EndEvent
	#tag Event
		Sub SelChange()
		  
		End Sub
	#tag EndEvent
	#tag Event
		Function KeyDown(Key As String) As Boolean
		  Return true
		End Function
	#tag EndEvent
#tag EndEvents
#tag Events TotalNumOfHosts_Data
	#tag Event
		Sub Open()
		  #IF TargetMacOS Then
		    Declare Sub setBezelStyle Lib "Cocoa" Selector "setBezelStyle:" (handle As Integer, value As Integer)
		    setBezelStyle(Me.Handle, 1)
		  #Endif
		End Sub
	#tag EndEvent
	#tag Event
		Function MouseDown(X As Integer, Y As Integer) As Boolean
		  Return True
		End Function
	#tag EndEvent
	#tag Event
		Function KeyDown(Key As String) As Boolean
		  Return True
		End Function
	#tag EndEvent
#tag EndEvents
#tag Events IANA_Class_Label_Data
	#tag Event
		Sub Open()
		  #IF TargetMacOS Then
		    Declare Sub setBezelStyle Lib "Cocoa" Selector "setBezelStyle:" (handle As Integer, value As Integer)
		    setBezelStyle(Me.Handle, 1)
		  #Endif
		End Sub
	#tag EndEvent
	#tag Event
		Function MouseDown(X As Integer, Y As Integer) As Boolean
		  REturn True
		End Function
	#tag EndEvent
	#tag Event
		Function KeyDown(Key As String) As Boolean
		  Return True
		End Function
	#tag EndEvent
#tag EndEvents
#tag Events SubnetMaskPrefix_Label_Data
	#tag Event
		Sub Open()
		  #IF TargetMacOS Then
		    Declare Sub setBezelStyle Lib "Cocoa" Selector "setBezelStyle:" (handle As Integer, value As Integer)
		    setBezelStyle(Me.Handle, 1)
		  #Endif
		End Sub
	#tag EndEvent
	#tag Event
		Function MouseDown(X As Integer, Y As Integer) As Boolean
		  Return True
		End Function
	#tag EndEvent
	#tag Event
		Function KeyDown(Key As String) As Boolean
		  Return True
		End Function
	#tag EndEvent
#tag EndEvents
#tag Events CiscoWildcardMask_Data_Label
	#tag Event
		Sub Open()
		  #IF TargetMacOS Then
		    Declare Sub setBezelStyle Lib "Cocoa" Selector "setBezelStyle:" (handle As Integer, value As Integer)
		    setBezelStyle(Me.Handle, 1)
		  #Endif
		End Sub
	#tag EndEvent
	#tag Event
		Function MouseDown(X As Integer, Y As Integer) As Boolean
		  Return True
		End Function
	#tag EndEvent
	#tag Event
		Function KeyDown(Key As String) As Boolean
		  Return True
		End Function
	#tag EndEvent
#tag EndEvents
#tag ViewBehavior
	#tag ViewProperty
		Name="AcceptFocus"
		Visible=true
		Group="Behavior"
		InitialValue="False"
		Type="Boolean"
		EditorType="Boolean"
	#tag EndViewProperty
	#tag ViewProperty
		Name="AcceptTabs"
		Visible=true
		Group="Behavior"
		InitialValue="True"
		Type="Boolean"
		EditorType="Boolean"
	#tag EndViewProperty
	#tag ViewProperty
		Name="AutoDeactivate"
		Visible=true
		Group="Appearance"
		InitialValue="True"
		Type="Boolean"
	#tag EndViewProperty
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
		Name="BeginOnClassLessSubnet"
		Group="Behavior"
		InitialValue="False"
		Type="Boolean"
	#tag EndViewProperty
	#tag ViewProperty
		Name="Enabled"
		Visible=true
		Group="Appearance"
		InitialValue="True"
		Type="Boolean"
		EditorType="Boolean"
	#tag EndViewProperty
	#tag ViewProperty
		Name="EraseBackground"
		Visible=true
		Group="Behavior"
		InitialValue="True"
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
		InitialValue="300"
		Type="Integer"
	#tag EndViewProperty
	#tag ViewProperty
		Name="HelpTag"
		Visible=true
		Group="Appearance"
		Type="String"
	#tag EndViewProperty
	#tag ViewProperty
		Name="InitialParent"
		Group="Position"
		Type="String"
	#tag EndViewProperty
	#tag ViewProperty
		Name="Left"
		Visible=true
		Group="Position"
		Type="Integer"
	#tag EndViewProperty
	#tag ViewProperty
		Name="LockBottom"
		Visible=true
		Group="Position"
		Type="Boolean"
	#tag EndViewProperty
	#tag ViewProperty
		Name="LockLeft"
		Visible=true
		Group="Position"
		Type="Boolean"
	#tag EndViewProperty
	#tag ViewProperty
		Name="LockRight"
		Visible=true
		Group="Position"
		Type="Boolean"
	#tag EndViewProperty
	#tag ViewProperty
		Name="LockTop"
		Visible=true
		Group="Position"
		Type="Boolean"
	#tag EndViewProperty
	#tag ViewProperty
		Name="Name"
		Visible=true
		Group="ID"
		Type="String"
	#tag EndViewProperty
	#tag ViewProperty
		Name="Super"
		Visible=true
		Group="ID"
		Type="String"
	#tag EndViewProperty
	#tag ViewProperty
		Name="TabIndex"
		Visible=true
		Group="Position"
		InitialValue="0"
		Type="Integer"
	#tag EndViewProperty
	#tag ViewProperty
		Name="TabPanelIndex"
		Group="Position"
		InitialValue="0"
		Type="Integer"
	#tag EndViewProperty
	#tag ViewProperty
		Name="TabStop"
		Visible=true
		Group="Position"
		InitialValue="True"
		Type="Boolean"
		EditorType="Boolean"
	#tag EndViewProperty
	#tag ViewProperty
		Name="Top"
		Visible=true
		Group="Position"
		Type="Integer"
	#tag EndViewProperty
	#tag ViewProperty
		Name="Transparent"
		Visible=true
		Group="Behavior"
		InitialValue="True"
		Type="Boolean"
		EditorType="Boolean"
	#tag EndViewProperty
	#tag ViewProperty
		Name="UseFocusRing"
		Visible=true
		Group="Appearance"
		InitialValue="False"
		Type="Boolean"
		EditorType="Boolean"
	#tag EndViewProperty
	#tag ViewProperty
		Name="User_UseSingleRangeOnly"
		Group="Behavior"
		InitialValue="False"
		Type="Boolean"
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
		InitialValue="300"
		Type="Integer"
	#tag EndViewProperty
#tag EndViewBehavior
