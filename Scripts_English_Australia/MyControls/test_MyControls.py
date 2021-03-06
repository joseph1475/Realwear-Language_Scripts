from playsound import playsound
import time, datetime
import os
import subprocess
import xlsxwriter
import re


TC_Name = "MyControls"
cwd = os.getcwd()
reports = r''+cwd+'/Reports/'+TC_Name+''
if not os.path.exists(reports):
    os.makedirs(reports)


screenshots = r''+cwd+'/Screenshots/'+TC_Name+''
if not os.path.exists(screenshots):
    os.makedirs(screenshots)

row = 0
col = 0
ts_report = time.time()
time_stamp_report = datetime.datetime.fromtimestamp(ts_report).strftime('-%d-%m-%Y-%H-%M-%S')
workbook = xlsxwriter.Workbook("Reports/"+TC_Name+"/"+TC_Name+""+time_stamp_report+".xlsx")
worksheet = workbook.add_worksheet()
worksheet.write(row, col, "Time_Stamp")
worksheet.write(row, col + 1, "Platform")
worksheet.write(row, col + 2, "BSP_Version")
worksheet.write(row, col + 3, "Expected")
worksheet.write(row, col + 4, "Actual")
worksheet.write(row, col + 5, "Confidence_Value")
worksheet.write(row, col + 6, "Result")
row += 1


def voice_commands(text):
    global row
    global col
    adb_clear = "adb shell logcat -c"
    os.system(adb_clear)
    adb_devinfo = "adb shell getprop ro.product.version.software"
    adb_process = subprocess.Popen(adb_devinfo, shell=True, stdout=subprocess.PIPE, stderr=subprocess.STDOUT)
    adb_bspversion = adb_process.communicate()[0]
    adb_bspversion = adb_bspversion.decode("utf-8")
    adb_platform = re.search(r'^\w+', str(adb_bspversion)).group(0)

    try:
        playsound('Tones/'+text+'.wav')
        time.sleep(1)
    except:
        playsound('Tones/'+text+'.mp3')
        time.sleep(1)
    ts = time.time()
    time_stamp = datetime.datetime.fromtimestamp(ts).strftime('-%d-%m-%Y-%H-%M-%S')
    adb_screenshot = "adb exec-out screencap -p > Screenshots/\""+TC_Name+"\"/\""+text+"\""+time_stamp+".png"
    os.system(adb_screenshot)
    adb_fetch = "adb shell logcat -d | find \"(ACCEPTED)\""
    adb_ps = subprocess.Popen(adb_fetch, shell=True, stdout=subprocess.PIPE, stderr=subprocess.STDOUT)
    adb_result = adb_ps.communicate()[0]
    adb_result = adb_result.decode("utf-8")
    try:
        adb_confidence = re.search(r"\[(.*?)\]", adb_result).group(1)
        adb_timestamp = re.search(r"\d{2}-\d{2} \d{2}:\d{2}:\d{2}.\d{3}", adb_result).group(0)
        adb_result = re.search(r"\)(.*?)\[", adb_result).group(1)
    except AttributeError:
        adb_result = ""
        adb_confidence = ""
        adb_timestamp = ""
    adb_result = adb_result[1:-1]
    if text.lower() == adb_result.lower():
        print(""+text+" -> PASS")
        assert adb_result.lower() == text.lower()
        worksheet.write(row, col, adb_timestamp)
        worksheet.write(row, col + 1, adb_platform)
        worksheet.write(row, col + 2, adb_bspversion)
        worksheet.write(row, col + 3, text)
        worksheet.write(row, col + 4, adb_result)
        worksheet.write(row, col + 5, adb_confidence)
        worksheet.write(row, col + 6, "PASS")
        row += 1
    else:
        print(""+text+" -> FAIL")
        worksheet.write(row, col, adb_timestamp)
        worksheet.write(row, col + 1, adb_platform)
        worksheet.write(row, col + 2, adb_bspversion)
        worksheet.write(row, col + 3, text)
        worksheet.write(row, col + 4, adb_result)
        worksheet.write(row, col + 5, adb_confidence)
        worksheet.write(row, col + 6, "FAIL")
        row += 1
        workbook.close()  # will save the result till it executed in current script
        assert adb_result.lower() == text.lower()  # will terminate execution if failure occurs


def close_recent_applications():
    try:
        playsound('Tones/Recent Applications.wav')
        time.sleep(3)
        playsound('Tones/Dismiss All.wav')
        time.sleep(2)
        playsound('Tones/Navigate Back.wav')
    except:
        playsound('Tones/Recent Applications.mp3')
        time.sleep(3)
        playsound('Tones/Dismiss All.mp3')
        time.sleep(2)
        playsound('Tones/Navigate Back.mp3')


def device_wakeup():
    adb_wakeup = "adb shell input keyevent KEYCODE_WAKEUP"
    os.system(adb_wakeup)

def test_MyControls():
    device_wakeup()
    close_recent_applications()
    time.sleep(2)

    voice_commands("Navigate Home")
    time.sleep(2)
    voice_commands("My Controls")
    time.sleep(2)
    voice_commands("Mouse")
    time.sleep(2)
    voice_commands("Auto Rotate")
    time.sleep(2)

    voice_commands("Action Button")
    time.sleep(1)
    voice_commands("Noise Capture")
    time.sleep(2)
    voice_commands("Action Button")
    time.sleep(1)
    voice_commands("Homescreen")
    time.sleep(2)
    voice_commands("Power Options")
    time.sleep(4)

    voice_commands("My Controls")
    time.sleep(2)
    voice_commands("WearML Indicators")
    time.sleep(1)
    voice_commands("Fade 5 Seconds")
    time.sleep(2)
    voice_commands("WearML Indicators")
    time.sleep(1)
    voice_commands("Fade 10 Seconds")
    time.sleep(2)
    voice_commands("WearML Indicators")
    time.sleep(1)
    voice_commands("Fade never")
    time.sleep(2)
    voice_commands("WearML Indicators")
    time.sleep(1)
    voice_commands("Fade 3 Seconds")
    time.sleep(4)
    voice_commands("Navigate Home")
    time.sleep(2)

    voice_commands("My Controls")
    time.sleep(2)
    voice_commands("Bluetooth")
    time.sleep(1)
    """
    voice_commands("Enable")
    time.sleep(5)
    voice_commands("Bluetooth")
    time.sleep(1)
    voice_commands("Disable")
    time.sleep(5)
    voice_commands("Bluetooth")
    time.sleep(1)
    """
    voice_commands("Bluetooth settings")
    time.sleep(3)
    voice_commands("Navigate Home")
    time.sleep(2)

    voice_commands("My Controls")
    time.sleep(3)
    voice_commands("Wireless Network")
    time.sleep(1)
    """
    voice_commands("Disable")
    time.sleep(5)
    voice_commands("Wireless Network")
    time.sleep(1)
    voice_commands("Enable")
    time.sleep(5)
    voice_commands("Wireless Network")
    time.sleep(1)
    """
    voice_commands("Wireless Network settings")
    time.sleep(5)
    voice_commands("Navigate Back")
    time.sleep(2)

    voice_commands("Navigate Home")
    time.sleep(2)
    voice_commands("My Controls")
    time.sleep(2)
    voice_commands("Preferred Network")
    time.sleep(2)
    voice_commands("Navigate Back")
    time.sleep(3)
    #flashlight
    voice_commands("Navigate Home")
    time.sleep(2)
    voice_commands("My Controls")
    time.sleep(2)
    voice_commands("Flashlight")
    time.sleep(2)
    voice_commands("Flashlight")
    time.sleep(2)
    voice_commands("Navigate Back")
    time.sleep(2)

    #Brightness
    voice_commands("My Controls")
    time.sleep(2)
    voice_commands("Brightness")
    time.sleep(1)
    voice_commands("Set level 1")
    time.sleep(5)
    voice_commands("Brightness")
    time.sleep(1)
    voice_commands("Set level 2")
    time.sleep(5)
    voice_commands("Brightness")
    time.sleep(1)
    voice_commands("Set level 3")
    time.sleep(5)
    voice_commands("Brightness")
    time.sleep(1)
    voice_commands("Set level 5")
    time.sleep(5)
    voice_commands("Brightness")
    time.sleep(1)
    voice_commands("Set level 6")
    time.sleep(5)
    voice_commands("Brightness")
    time.sleep(1)
    voice_commands("Set level 7")
    time.sleep(5)
    voice_commands("Brightness")
    time.sleep(1)
    voice_commands("Set level 8")
    time.sleep(5)
    voice_commands("Brightness")
    time.sleep(1)
    voice_commands("Set level 9")
    time.sleep(5)
    voice_commands("Brightness")
    time.sleep(1)
    voice_commands("Set level 10")
    time.sleep(5)
    voice_commands("Brightness")
    time.sleep(1)
    voice_commands("Set level 4")
    time.sleep(5)
    voice_commands("Navigate Home")
    time.sleep(2)

    #Volume
    voice_commands("My Controls")
    time.sleep(2)
    voice_commands("Volume")
    time.sleep(1)
    voice_commands("Set level 1")
    time.sleep(5)
    voice_commands("Volume")
    time.sleep(1)
    voice_commands("Set level 2")
    time.sleep(5)
    voice_commands("Volume")
    time.sleep(1)
    voice_commands("Set level 3")
    time.sleep(5)
    voice_commands("Volume")
    time.sleep(1)
    voice_commands("Set level 4")
    time.sleep(5)
    voice_commands("Volume")
    time.sleep(1)
    voice_commands("Set level 6")
    time.sleep(5)
    voice_commands("Volume")
    time.sleep(1)
    voice_commands("Set level 7")
    time.sleep(5)
    voice_commands("Volume")
    time.sleep(1)
    voice_commands("Set level 8")
    time.sleep(5)
    voice_commands("Volume")
    time.sleep(1)
    voice_commands("Set level 9")
    time.sleep(5)
    voice_commands("Volume")
    time.sleep(1)
    voice_commands("Set level 10")
    time.sleep(5)
    voice_commands("Volume")
    time.sleep(1)
    voice_commands("Set level 5")
    time.sleep(5)

    #voice_commands("Microphone")
    #time.sleep(2)
    #voice_commands("Microphone")
    #time.sleep(2)
    voice_commands("Color mode")
    time.sleep(2)
    voice_commands("Color mode")
    time.sleep(2)
    voice_commands("Dictation")
    time.sleep(2)
    voice_commands("Dictation")
    time.sleep(2)
    voice_commands("Help command")
    time.sleep(2)
    voice_commands("Help command")
    time.sleep(2)
    voice_commands("Layout mode")
    time.sleep(2)
    voice_commands("Page down")
    time.sleep(2)
    voice_commands("Layout mode")
    time.sleep(2)
    voice_commands("More settings")
    time.sleep(2)
    voice_commands("Navigate back")
    time.sleep(2)

    workbook.close()



