#  python serial2keyboard.py COM11 9600 "test.txt - Editor"

#pip3 install pynput pyserial



import serial
from pynput.keyboard import Key, Controller
import sys
import glob
import win32gui
import win32con
import win32api
from time import sleep


def list_ports():
	""" Finds all serial ports and returns a list containing them

		:raises EnvironmentError:
			On unsupported or unknown platforms
		:returns:
			A list of the serial ports available on the system
	"""
	if sys.platform.startswith('win'):
		ports = ['COM%s' % (i + 1) for i in range(256)]
	elif sys.platform.startswith('linux') or sys.platform.startswith('cygwin'):
		# this excludes your current terminal "/dev/tty"
		ports = glob.glob('/dev/tty[A-Za-z]*')
	elif sys.platform.startswith('darwin'):
		ports = glob.glob('/dev/tty.*')
	else:
		raise EnvironmentError('Unsupported platform')

	result = []
	for port in ports:
		try:
			s = serial.Serial(port)	# Try to open a port
			s.close()				  # Close the port if sucessful
			result.append(port)		# Add to list of good ports
		except (OSError, Exception):   # If un sucessful
			pass
	return result 


def winEnumHandler( hwnd, ctx ):
    if win32gui.IsWindowVisible( hwnd ):
        print (hex(hwnd), "'",win32gui.GetWindowText( hwnd ),"'")

hwndChild = None
sendBackground = False
if(len(sys.argv)>3):
	winName=sys.argv[3]  # "test.txt - Editor"
	hwndMain = win32gui.FindWindow(None, winName) 
	if 0 == hwndMain :
		hwndMain = win32gui.FindWindow(None, "*" + winName)  
	if 0 != hwndMain :
		if sendBackground : 
			hwndChild = win32gui.GetWindow(hwndMain, win32con.GW_CHILD)
		else :
			win32gui.SetForegroundWindow(hwndMain)  # show that i found it ... 
	print ("searched for win \""+winName+"\" " + str(hwndMain) + " " + str(hwndChild))




# https://docs.microsoft.com/de-de/windows/win32/inputdev/virtual-key-codes
# http://timgolden.me.uk/pywin32-docs/contents.html
print(win32gui.SetForegroundWindow(hwndMain))  # show that i found it ... 
##### FAIL
#print(win32api.SendMessage(hwndChild, win32con.WM_KEYDOWN, 0x12, None)) # win32con.VK_MENU ALT --> FAIL
#sleep(1)
#print(win32api.SendMessage(hwndChild, win32con.WM_KEYUP, 0x5A, None)) 

#win32api.PostMessage(hwndChild, win32con.WM_CHAR, 0x12, 0) # ALT
#win32api.PostMessage(hwndChild, win32con.WM_CHAR, ord("o"), 0) # o
#win32api.PostMessage(hwndChild, win32con.WM_CHAR, ord("z"), 0) # z



if (len(sys.argv)<2  or  ( len(sys.argv)>3 and 0 == hwndMain) ) :
	print("usage:")
	print("To run the python script: 'python3 serial2keyboard.py <Port> [baud rate [WindowName]] ', eg: 'python3 serial2keyboard.py /dev/ttyUSB0', or 'python3 serial2keyboard.py /dev/ttyUSB0 9600'")
	print("To run the native executable made with pyinstaller: './serial2keyboard(probably add .exe on windows) [Port] [optional: baudrate]'")
	print("eg.: python serial2keyboard.py 11 9600 \"test.txt - Editor\"")
	print("Default baud rate if not supplied as second argument: 9600")
	print("Suspected suitable serial ports are:", list_ports())
	if  len(sys.argv)>3  :
		print("Suspected suitable windows:")
		win32gui.EnumWindows( winEnumHandler, None )
	sys.exit(1)


print("Attempting to open serial port: ", sys.argv[1])
baudRate=9600
if(len(sys.argv)>2):
	baudRate=int(sys.argv[2])
print("Setting baud rate to",baudRate)

try:
	s = serial.Serial(sys.argv[1],baudRate,timeout=1)
except:
	print("Failed to open", sys.argv[1], "as a serial port.. are you doing this right?")
	sys.exit(1)

print("Shit, it worked.. waiting for serial data...")

keyboard = Controller()




if 1==1 :
#try:
	while(s.is_open):

		if(s.in_waiting>0):
			rxLine=s.readline().decode("ascii").strip()
			print(rxLine)
			if False and 0 != hwndChild :
				print("hwndChild")
			    #[hwndChild] this is the Unique ID of the sub/child application/proccess
			    #[win32con.WM_CHAR] This sets what PostMessage Expects for input theres KeyDown and KeyUp as well
			    #[0x44] hex code for D
			    #[0]No clue, good luck!
			    #
			    #temp = win32api.PostMessage(hwndChild, win32con.WM_CHAR, 0x44, 0) returns key sent
		    	#print(temp)
		    	#
		    	#
		    	# !!!!!!!!!!!!!!!! setzt text - ersetzt GESAMTEN inhalt
				#win32api.SendMessage(hwndChild, win32con.WM_SETTEXT, None, rxLine)
				#win32api.SendMessage(hwndChild, win32con.WM_KEYDOWN, win32con.VK_RETURN, None)
				#win32api.SendMessage(hwndChild, win32con.WM_KEYUP, win32con.VK_RETURN, None)
				#
				#win32api.PostMessage(hwndChild, win32con.WM_SETTEXT, None, rxLine) #  pywintypes.error: (1159, 'PostMessage', 'Diese Nachricht kann nur mit synchronen Vorgängen verwendet werden.')
				for c in rxLine:
					temp = win32api.PostMessage(hwndChild, win32con.WM_CHAR, ord(c), 0) # returns key sent
				temp = win32api.PostMessage(hwndChild, win32con.WM_CHAR, 0x0D, 0) # returns key sent
				#win32api.SendMessage(hwndChild, win32con.WM_KEYDOWN, win32con.VK_RETURN, 0)
				#win32api.SendMessage(hwndChild, win32con.WM_KEYUP, win32con.VK_RETURN, 0)
			else:
				if 0 != hwndMain :
					win32gui.SetForegroundWindow(hwndMain)
				if "Mute" == rxLine : # Ctrl + Shift + M
					keyboard.press(Key.cmd.value)
					keyboard.press(Key.shift.value)
					keyboard.press('m')
					keyboard.release('m')
					keyboard.release(Key.shift.value)
					keyboard.release(Key.cmd.value)
				elif "Cam" == rxLine :
					keyboard.press(Key.cmd.value)
					keyboard.press(Key.shift.value)
					keyboard.press('o')
					keyboard.release('o')
					keyboard.release(Key.shift.value)
					keyboard.release(Key.cmd.value)
				else:
					keyboard.type(rxLine)
					keyboard.press(Key.enter)
					keyboard.release(Key.enter)
#except:
#	print("Something happened... did you just yank the thing out?")