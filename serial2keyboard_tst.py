#you will need the win32 libraries for this snippet of code to work, Links below
import win32gui
import win32con
import win32api
from time import sleep



def winEnumHandler( hwnd, ctx ):
    if win32gui.IsWindowVisible( hwnd ):
        print (hex(hwnd), "'",win32gui.GetWindowText( hwnd ),"'")




#[hwnd] No matter what people tell you, this is the handle meaning unique ID, 
#["Notepad"] This is the application main/parent name, an easy way to check for examples is in Task Manager
#["test - Notepad"] This is the application sub/child name, an easy way to check for examples is in Task Manager clicking dropdown arrow
#hwndMain = win32gui.FindWindow("Notepad", "test - Notepad") this returns the main/parent Unique ID
#hwndMain = win32gui.FindWindow("Notepad", "test - Notepad")
#hwndMain = win32gui.FindWindow("Editor", "test.txt - Editor")
#hwndMain = win32gui.FindWindow("Editor", "test - Editor")
#hwndMain = win32gui.FindWindow("Editor", None)
#hwndMain = win32gui.FindWindow("Notepad", None) # even when taskmanager names Editor
#hwndMain = win32gui.FindWindow("Notepad", "test.txt - Notepad") # 
#hwndMain = win32gui.FindWindow("Notepad", "test.txt - Editor") # OK
hwndMain = win32gui.FindWindow(None, "test.txt - Editor") # OK
if 0 == hwndMain :
	hwndMain = win32gui.FindWindow("Notepad", "*test.txt - Editor") # OK
if 0 == hwndMain :
	win32gui.EnumWindows( winEnumHandler, None )


win32gui.SetForegroundWindow(hwndMain)


#["hwndMain"] this is the main/parent Unique ID used to get the sub/child Unique ID
#[win32con.GW_CHILD] I havent tested it full, but this DOES get a sub/child Unique ID, if there are multiple you'd have too loop through it, or look for other documention, or i may edit this at some point ;)
#hwndChild = win32gui.GetWindow(hwndMain, win32con.GW_CHILD) this returns the sub/child Unique ID
hwndChild = win32gui.GetWindow(hwndMain, win32con.GW_CHILD)

#print(hwndMain) #you can use this to see main/parent Unique ID
#print(hwndChild)  #you can use this to see sub/child Unique ID

#while(True):

    #[hwndChild] this is the Unique ID of the sub/child application/proccess
    #[win32con.WM_CHAR] This sets what PostMessage Expects for input theres KeyDown and KeyUp as well
    #[0x44] hex code for D
    #[0]No clue, good luck!
    #temp = win32api.PostMessage(hwndChild, win32con.WM_CHAR, 0x44, 0) returns key sent

	#temp = win32api.PostMessage(hwndChild, win32con.WM_CHAR, 0x44, 0)
	#print(temp)
	#sleep(1)