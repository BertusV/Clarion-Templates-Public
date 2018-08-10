#Template(TileButton,'Tile Button Class Templates'),Family('ABC')
#!------------------------------------------------------------------------------------!#
#! Tile Button Version 1.0
#! (c) 2018 by Bertus Viljoen
#! All rights reserved
#! This Template is written to implement the Tile Button class written by 
#! Brahn Partridge running ClarionHub.com
#!
#!------------------------------------------------------------------------------------!#
#Extension(Activate_TileButton,'Activate Tile Button'),Application
  #Display('Tile Button')
  #Display('(c) 2018 All Rights Reserved')
  #Display('Version 1.0')
  #Display()
  #Display('See "Adding Tile Button to your application."')
  #Display('This template implements Brahn''s Tile Button Class.')
  #Display()
  #Prompt('Disable Tile Button in this App',Check),%DisableTB,AT(10)
#!
#!------------------------------------------------------------------------------------!#
#AT (%AfterGlobalIncludes),where(%DisableTB=0)
  #Call(%TileButtonInclude)
#ENDAT
#!------------------------------------------------------------------------------------!#
#!

#Extension(IncludeTileButtonManager,'Tile Button Manager'),Description('Tile Button Manager'),Procedure,Req(Activate_TileButton)
#Sheet
#Tab('General')
#Boxed('Tile Button')
  #Prompt('Tile Manager object name',@S25),%TileManager,DEFAULT('Tiles'),REQ
  #Prompt('Tile Buttons',Control),%TileButtons,Multi('Tile Buttons'),REQ
  #Prompt('Disable Tile Button in this Procedure',Check),%DisableTBP,AT(10)
#EndBoxed
#EndTab
#EndSheet

#! Declare Local data
#AT (%DataSection,''),where(%DisableTB=%False AND %DisableTBP=%False),Priority(8500)
     #Call(%TileButtonLocalData)
#EndAt

     
#AT (%WindowManagerMethodCodeSection,'Init'),where(%DisableTB=%False AND %DisableTBP=%False),Priority(8010)
#Call(%TileButtonInitiate)
#FOR(%TileButtons)
     #Call(%TileButtonAssignment)
#ENDFOR
#ENDAT


#!
#!------------------------------------------------------------------------------------!#
#!                                                                                   		        !#
#! This Groups are part of the Tile Button template and Classes      	        !#
#!                                                                                    		        !#
#!------------------------------------------------------------------------------------!#

#Group(%TileButtonInclude),Preserve
#If (%DisableTB=%False)
    Include('TileManager.inc'),ONCE
#EndIf

#Group(%TileButtonLocalData),Preserve
#If (%DisableTB=%False)
TM:%TileManager                        TileManager
#EndIf

#Group(%TileButtonAssignment),Preserve
#If (%DisableTB=%False)
        TM:%TileManager.AddButtonMimic(%TileButtons, %TileButtons{PROP:Background})
#EndIf

#Group(%TileButtonInitiate),Preserve
#If (%DisableTB=%False)
        TM:%TileManager.Init(SELF)
#EndIf
        
