from playsound import playsound
import time, datetime
import os
import subprocess
import xlsxwriter
import re


TC_Name = "MyFiles"
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

def test_MyFiles():
    device_wakeup()
    close_recent_applications()
    time.sleep(2)

    #first level UI
    voice_commands("Navigate Home")
    time.sleep(2)
    voice_commands("My Files")
    time.sleep(2)
    voice_commands("My Media")
    time.sleep(2)
    voice_commands("Navigate back")
    time.sleep(2)
    voice_commands("My Documents")
    time.sleep(2)
    voice_commands("Navigate back")
    time.sleep(2)
    voice_commands("My Photos")
    time.sleep(2)
    voice_commands("Navigate back")
    time.sleep(2)
    voice_commands("My Downloads")
    time.sleep(2)
    voice_commands("Navigate back")
    time.sleep(2)
    voice_commands("My Drives")
    time.sleep(2)
    voice_commands("My device")
    time.sleep(2)
    voice_commands("Alarms")
    time.sleep(2)
    voice_commands("Navigate back")
    time.sleep(2)
    voice_commands("DCIM")
    time.sleep(2)
    voice_commands("Navigate back")
    time.sleep(2)
    voice_commands("Documents")
    time.sleep(2)
    voice_commands("Navigate back")
    time.sleep(2)
    voice_commands("Download")
    time.sleep(2)
    voice_commands("Navigate back")
    time.sleep(2)
    voice_commands("Manage")
    time.sleep(2)
    voice_commands("Navigate back")
    time.sleep(2)
    voice_commands("Movies")
    time.sleep(2)
    voice_commands("Navigate back")
    time.sleep(2)
    #voice_commands("Music")
    #time.sleep(2)
    #voice_commands("Navigate back")
    #time.sleep(2)
    voice_commands("Notifications")
    time.sleep(2)
    voice_commands("Navigate back")
    time.sleep(2)
    voice_commands("Pictures")
    time.sleep(2)
    voice_commands("Back One level")
    time.sleep(2)
    voice_commands("Podcasts")
    time.sleep(2)
    voice_commands("Back One level")
    time.sleep(2)
    voice_commands("Realwear")
    time.sleep(2)
    voice_commands("Navigate back")
    time.sleep(2)
    voice_commands("Ringtones")
    time.sleep(2)
    voice_commands("Recent Applications")
    time.sleep(3)
    voice_commands("Dismiss All")
    time.sleep(2)

    #Verify to Open My Photos
    voice_commands("Navigate HOME")
    time.sleep(2)
    voice_commands("My CAMERA")
    time.sleep(2)
    voice_commands("Take Photo")
    time.sleep(5)
    voice_commands("Navigate Home")
    time.sleep(2)
    voice_commands("MY FILES")
    time.sleep(2)
    voice_commands("MY PHOTOS")
    time.sleep(2)
    voice_commands("CAMERA")
    time.sleep(2)
    voice_commands("SELECT ITEM 1")
    time.sleep(2)
    voice_commands("ZOOM LEVEL 2")
    time.sleep(2)
    voice_commands("ZOOM LEVEL 1")
    time.sleep(2)
    voice_commands("ZOOM LEVEL 4")
    time.sleep(2)
    voice_commands("ZOOM LEVEL 5")
    time.sleep(2)
    voice_commands("ZOOM LEVEL 3")
    time.sleep(2)
    voice_commands("FREEZE WINDOW")
    time.sleep(2)
    voice_commands("CONTROL WINDOW")
    time.sleep(2)
    voice_commands("RESET WINDOW")
    time.sleep(2)
    voice_commands("Navigate back")
    time.sleep(2)
    voice_commands("Back one level")
    time.sleep(2)
    voice_commands("MORE OPTIONS")
    time.sleep(2)
    voice_commands("SORT FILES")
    time.sleep(2)
    voice_commands("BY DATE")
    time.sleep(2)
    voice_commands("Sort Files")
    time.sleep(2)
    voice_commands("By Name")
    time.sleep(2)
    voice_commands("Hide Options")
    time.sleep(2)
    voice_commands("Camera")
    time.sleep(2)
    voice_commands("MORE OPTIONS")
    time.sleep(2)
    voice_commands("EDIT MODE")
    time.sleep(2)
    voice_commands("SELECT ITEM 1")
    time.sleep(2)
    voice_commands("DELETE SELECTED")
    time.sleep(2)
    voice_commands("CONFIRM DELETION")
    time.sleep(2)
    voice_commands("Recent Applications")
    time.sleep(2)
    voice_commands("Dismiss All")
    time.sleep(2)
    #Verify select item in My Documents
    voice_commands("MY FILES")
    time.sleep(2)
    voice_commands("MY DOCUMENTS")
    time.sleep(2)
    voice_commands("SELECT ITEM 1")
    time.sleep(5)
    voice_commands("ZOOM LEVEL 3")
    time.sleep(2)
    voice_commands("ZOOM LEVEL 1")
    time.sleep(2)
    voice_commands("ZOOM LEVEL 4")
    time.sleep(2)
    voice_commands("ZOOM LEVEL 5")
    time.sleep(2)
    voice_commands("ZOOM LEVEL 2")
    time.sleep(2)
    voice_commands("FREEZE WINDOW")
    time.sleep(2)
    voice_commands("CONTROL WINDOW")
    time.sleep(2)
    voice_commands("Reset window")
    time.sleep(2)
    voice_commands("GO TO PAGE 10")
    time.sleep(2)
    voice_commands("NEXT PAGE")
    time.sleep(2)
    voice_commands("PREVIOUS PAGE")
    time.sleep(2)
    workbook.close()
