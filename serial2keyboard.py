#pip3 install pynput pyserial

import serial
from pynput.keyboard import Key, Controller
import sys
import glob
import win32gui
import win32con
import win32api


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

if(len(sys.argv)>3):
	winName=sys.argv[3]  # "test.txt - Editor"
	hwndMain = win32gui.FindWindow(None, winName) 
	if 0 == hwndMain :
		hwndMain = win32gui.FindWindow(None, "*" + winName)  
	if 0 != hwndMain :
		win32gui.SetForegroundWindow(hwndMain)  # show that i found it ... 
	print ("searched for win \""+winName+"\"")

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

try:
	while(s.is_open):

		if(s.in_waiting>0):
			if 0 != hwndMain :
				win32gui.SetForegroundWindow(hwndMain)
			rxLine=s.readline().decode("ascii").strip()
			keyboard.type(rxLine)
			keyboard.press(Key.enter)
			keyboard.release(Key.enter)
except:
	print("Something happened... did you just yank the thing out?")