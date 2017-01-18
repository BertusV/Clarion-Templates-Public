#TEMPLATE(AltKeyFix,'Bertus Viljoen'),FAMILY('ABC'),FAMILY('CW20')
#Extension(AltKeyFix,'Bertus Viljoen Alt key menu Fix'),APPLICATION
#!--------------------------------------------------------------------------
#!
#! Author: Bertus Viljoen
#!
#! Original Author: Marius van den Berg
#!
#! Fixes the bug where the ALT button hangs an application
#! If a window is open in an application in windows, the app will freeze up if the user presses ALT key
#! to access the file menu
#!
#! The problem occurs if the ALT/F10 was only pressed without any hot keys
#! So basically all the template does is, if only ALT was pressed when any windows
#! in the app are open I capture the KeyCode - event and do nothing
#!
#! If no windows are open however, the app will call the menupress manually.
#!
#! Note:  Notes indicate this problem is solved in Clarion 7.1 but recheck at that point - !MW 01/18/11
#! Note: This surfaced again in Windows 10/ Clarion 10.
#!----------------------------------------------------------------------------------------------------
#Sheet
	#Tab('General')
		#BOXED
			#Display('Bertus Viljoen')
			#Display('Windows ALT Button Fix')
		#ENDBOXED
		#BOXED('Debugging')
			#PROMPT('Disable ALT button fix',Check),%NoGloALTButtonFix,At(10)
		#ENDBOXED
	#EndTab
#EndSheet
#!----------------------------------------------------------------------------------------------------
#! Set the altkey as an alertkey
#!#At(%DataSection,'WinFix')
#!ALT:Flag        Byte
#!#EndAt
#!----------------------------------------------------------------------------------------------------
#AT(%AfterWindowOpening),Where(%NoGloALTButtonFix=0)
   #If(%WindowStatement)
    #!If(%MenuBarStatement)
   ! Set the Tab Key And F10 Key as Alert Keys
   %Window{Prop:Alrt,200}=AltKeyPressed
   %Window{Prop:Alrt,201}=F10Key
    #!EndIf
   #EndIf
#ENDAT
#!----------------------------------------------------------------------------------------------------
#AT(%WindowEventHandling,'AlertKey'),priority(6000),Where(%NoGloALTButtonFix=0)
   #If(%WindowStateMent)
    #!If(%MenuBarStateMent)
   ! Only select the menu if no window was open
   If KeyCode() = AltKeyPressed ! If Tab was Pressed
      If Thread() = 1 Then
        PressKey(AltKeyPressed)
      End
   End
   If KeyCode() = F10Key ! If F10 was Pressed
      If Thread() = 1 Then
        PressKey(F10Key)
      End
   End
    #!EndIf
   #EndIf
#ENDAT
#!----------------------------------------------------------------------------------------------------

		