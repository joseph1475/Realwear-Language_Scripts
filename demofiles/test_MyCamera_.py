from playsound import playsound
import time, datetime
import os
import subprocess
import xlsxwriter
import re


TC_Name = "MyCamera_French"
cwd = os.getcwd()
reports = r''+cwd+'/Reports/'+TC_Name+''
if not os.path.exists(reports):
    os.makedirs(reports)

ts_screenshot = time.time()
time_stamp_screenshot = datetime.datetime.fromtimestamp(ts_screenshot).strftime('-%d-%m-%Y-%H-%M-%S')
screenshots = r''+cwd+'/Screenshots/'+TC_Name+''+time_stamp_screenshot+''
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
worksheet.write(row, col + 3, "English_Command")
worksheet.write(row, col + 4, "Expected_Result")
worksheet.write(row, col + 5, "Actual_Result")
worksheet.write(row, col + 6, "Confidence_Value")
worksheet.write(row, col + 7, "Result")
row += 1


def voice_commands(text, text_log):
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
    adb_screenshot = "adb exec-out screencap -p > Screenshots/\""+TC_Name+""+time_stamp_screenshot+"\"/\""+text+"\""+time_stamp+".png"
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

    if text_log.lower() == adb_result.lower():
        print(""+text_log+" -> PASS")
        worksheet.write(row, col, adb_timestamp)
        worksheet.write(row, col + 1, adb_platform)
        worksheet.write(row, col + 2, adb_bspversion)
        worksheet.write(row, col + 3, text)
        worksheet.write(row, col + 4, text_log)
        worksheet.write(row, col + 5, adb_result)
        worksheet.write(row, col + 6, adb_confidence)
        worksheet.write(row, col + 7, "PASS")
        row += 1
    else:
        print(""+text_log+" -> FAIL")
        worksheet.write(row, col, adb_timestamp)
        worksheet.write(row, col + 1, adb_platform)
        worksheet.write(row, col + 2, adb_bspversion)
        worksheet.write(row, col + 3, text)
        worksheet.write(row, col + 4, text_log)
        worksheet.write(row, col + 5, adb_result)
        worksheet.write(row, col + 6, adb_confidence)
        worksheet.write(row, col + 7, "FAIL")
        workbook.close()  # will save the result till it executed in current script
        assert adb_result.lower() == text.lower()  # will terminate execution if failure occurs
        row += 1


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

def test_MyCamera_French():
    device_wakeup()
    close_recent_applications()
    time.sleep(2)
    voice_commands("My Programs","")
    time.sleep(2)
    voice_commands("My Camera","")
    time.sleep(2)
    voice_commands("Take Photo","")
    time.sleep(5)
    try:
        voice_commands("Preview","")
        time.sleep(2)
        voice_commands("zoom level 2","")
        time.sleep(2)
        voice_commands("zoom level 3","")
        time.sleep(2)
        voice_commands("zoom level 4","")
        time.sleep(2)
        voice_commands("zoom level 5","")
        time.sleep(2)
        voice_commands("Freeze Window","")
        time.sleep(2)
        voice_commands("Control Window","")
        time.sleep(2)
        voice_commands("zoom level 1","")
        time.sleep(2)
        voice_commands("Delete","")
        time.sleep(2)
        voice_commands("Confirm","")
        time.sleep(2)
    finally:
        time.sleep(2)
        voice_commands("Recent Applications","")
        time.sleep(2)
        voice_commands("Dismiss All","")
        time.sleep(2)
        voice_commands("My Camera","")
        time.sleep(2)
        voice_commands("Take Photo","")
        time.sleep(10)
        voice_commands("My Files","")
        time.sleep(2)
        voice_commands("My Photos","")
        time.sleep(2)
        voice_commands("Camera","")
        time.sleep(2)
        voice_commands("Select item 1","")
        time.sleep(2)
        voice_commands("Navigate Back","")
        time.sleep(2)
        voice_commands("Back one level","")
        time.sleep(2)
        voice_commands("My files","")
        time.sleep(2)
        voice_commands("Navigate Back","")
        time.sleep(2)
        voice_commands("Show Help","")
        time.sleep(2)
        voice_commands("Hide Help","")
        time.sleep(2)
        voice_commands("Flash On","")
        time.sleep(2)
        voice_commands("Flash Off","")
        time.sleep(2)
        voice_commands("Flash Auto","")
        time.sleep(2)
        voice_commands("Manual Focus","")
        time.sleep(2)
        voice_commands("Auto Focus","")
        time.sleep(2)
        voice_commands("Exposure level plus 1","")
        time.sleep(2)
        voice_commands("Exposure level plus 2","")
        time.sleep(2)
        voice_commands("Exposure level minus 1","")
        time.sleep(2)
        voice_commands("Exposure level minus 2","")
        time.sleep(2)
        voice_commands("Exposure level 0","")
        time.sleep(2)
        voice_commands("zoom level 2","")
        time.sleep(2)
        voice_commands("zoom level 3","")
        time.sleep(2)
        voice_commands("zoom level 4","")
        time.sleep(2)
        voice_commands("zoom level 5","")
        time.sleep(2)
        voice_commands("zoom level 1","")
        time.sleep(2)
        voice_commands("More options","")
        time.sleep(2)
        voice_commands("Aspect ratio","")
        time.sleep(2)
        voice_commands("4 by 3","")
        time.sleep(2)
        voice_commands("Aspect ratio","")
        time.sleep(2)
        voice_commands("16 by 9","")
        time.sleep(2)
        voice_commands("Field of view","")
        time.sleep(2)
        voice_commands("Image Resolution","")
        time.sleep(2)
        voice_commands("Low","")
        time.sleep(2)
        voice_commands("Image Resolution","")
        time.sleep(2)
        voice_commands("High","")
        time.sleep(2)
        voice_commands("Video Resolution","")
        time.sleep(2)
        voice_commands("Low","")
        time.sleep(2)
        voice_commands("Video Resolution","")
        time.sleep(2)
        voice_commands("High","")
        time.sleep(2)
        voice_commands("Frame rate","")
        time.sleep(2)
        voice_commands("15","")
        time.sleep(2)
        voice_commands("Frame rate","")
        time.sleep(2)
        voice_commands("30","")
        time.sleep(2)
        voice_commands("Frame rate","")
        time.sleep(2)
        voice_commands("25","")
        time.sleep(2)
        voice_commands("Video Stabilization","")
        time.sleep(2)
        voice_commands("Hide Options","")
        time.sleep(2)
        voice_commands("Record Video","")
        time.sleep(10)
        voice_commands("Stop Recording","")
        time.sleep(2)
        voice_commands("Preview","")
        time.sleep(2)
        voice_commands("Video Pause","")
        time.sleep(2)
        voice_commands("Video Play","")
        time.sleep(2)
        voice_commands("Video Rewind","")
        time.sleep(2)
        voice_commands("Video Forward","")
        time.sleep(2)
        voice_commands("Video Stop","")
        time.sleep(2)
        voice_commands("Navigate Home","")
        time.sleep(2)

        workbook.close()


