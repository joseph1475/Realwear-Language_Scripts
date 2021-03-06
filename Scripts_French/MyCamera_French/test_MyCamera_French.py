from playsound import playsound
import time, datetime
import os
import subprocess
import xlsxwriter
import re

TC_Name = "MyCamera_French"
cwd = os.getcwd()
reports = r'' + cwd + '/Reports/' + TC_Name + ''
if not os.path.exists(reports):
    os.makedirs(reports)

ts_screenshot = time.time()
time_stamp_screenshot = datetime.datetime.fromtimestamp(ts_screenshot).strftime('-%d-%m-%Y-%H-%M-%S')
screenshots = r'' + cwd + '/Screenshots/' + TC_Name + '' + time_stamp_screenshot + ''
if not os.path.exists(screenshots):
    os.makedirs(screenshots)

row = 0
col = 0
ts_report = time.time()
time_stamp_report = datetime.datetime.fromtimestamp(ts_report).strftime('-%d-%m-%Y-%H-%M-%S')
workbook = xlsxwriter.Workbook("Reports/" + TC_Name + "/" + TC_Name + "" + time_stamp_report + ".xlsx")
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
        playsound('Tones/' + text + '.wav')
        time.sleep(1)
    except:
        playsound('Tones/' + text + '.mp3')
        time.sleep(1)
    ts = time.time()
    time_stamp = datetime.datetime.fromtimestamp(ts).strftime('-%d-%m-%Y-%H-%M-%S')
    adb_screenshot = "adb exec-out screencap -p > Screenshots/\"" + TC_Name + "" + time_stamp_screenshot + "\"/\"" + text + "\"" + time_stamp + ".png"
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
        print("" + text_log + " -> PASS")
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
        print("" + text_log + " -> FAIL")
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
    voice_commands("My Programs", "Mes programmes")
    time.sleep(2)
    voice_commands("My Camera", "Mon appareil photo")
    time.sleep(2)
    voice_commands("Take Photo", "Prendre une Photo")
    time.sleep(5)
    voice_commands("Preview", "APER??U")
    time.sleep(2)
    voice_commands("zoom level 2", "Niveau de zoom 2")
    time.sleep(2)
    voice_commands("zoom level 3", "Niveau de zoom 3")
    time.sleep(2)
    voice_commands("zoom level 4", "Niveau de zoom 4")
    time.sleep(2)
    voice_commands("zoom level 5", "Niveau de zoom 5")
    time.sleep(2)
    voice_commands("Freeze Window", "Gel ??cran")
    time.sleep(2)
    voice_commands("Control Window", "D??blocage ??cran")
    time.sleep(2)
    voice_commands("zoom level 1", "Niveau de zoom 1")
    time.sleep(2)
    voice_commands("Delete", "Effacer")
    time.sleep(2)
    voice_commands("Confirm", "CONFIRMER")
    time.sleep(2)
    time.sleep(2)
    voice_commands("Recent Applications", "Applications r??centes")
    time.sleep(2)
    voice_commands("Dismiss All", "Tout rejeter")
    time.sleep(2)
    voice_commands("My Camera", "Mon appareil photo")
    time.sleep(2)
    voice_commands("Take Photo", "Prendre une Photo")
    time.sleep(10)
    voice_commands("My Files", "Mes Fichiers")
    time.sleep(2)
    voice_commands("My Photos", "Mes Photos")
    time.sleep(2)
    voice_commands("Camera", "Camera")
    time.sleep(2)
    voice_commands("Select item 1", "S??lectionnez l'??l??ment 1")
    time.sleep(2)
    voice_commands("Navigate Back", "Retour En Arri??re")
    time.sleep(2)
    voice_commands("Back one level", "Retour d'un niveau")
    time.sleep(2)
    voice_commands("My files", "MES FICHIERS")
    time.sleep(2)
    voice_commands("Navigate Back", "Retour En Arri??re")
    time.sleep(2)
    voice_commands("Show Help", "Afficher l'aide")
    time.sleep(2)
    voice_commands("Hide Help", "Masquer l'aide")
    time.sleep(2)
    voice_commands("Flash On", "Lampe de poche allum??e")
    time.sleep(2)
    voice_commands("Flash Off", "Lampe de poche ??teinte")
    time.sleep(2)
    voice_commands("Flash Auto", "Flash automatique")
    time.sleep(2)
    voice_commands("Manual Focus", "Mise au point manuelle")
    time.sleep(2)
    voice_commands("Auto Focus", "Mise au point automatique")
    time.sleep(2)
    voice_commands("Exposure level plus 1", "Niveau d'exposition plus 1")
    time.sleep(2)
    voice_commands("Exposure level plus 2", "Niveau d'exposition plus 2")
    time.sleep(2)
    voice_commands("Exposure level minus 1", "Niveau d'exposition moins 1")
    time.sleep(2)
    voice_commands("Exposure level minus 2", "Niveau d'exposition moins 2")
    time.sleep(2)
    voice_commands("Exposure level 0", "Niveau d'exposition 0")
    time.sleep(2)
    voice_commands("zoom level 2", "Niveau de zoom 2")
    time.sleep(2)
    voice_commands("zoom level 3", "Niveau de zoom 3")
    time.sleep(2)
    voice_commands("zoom level 4", "Niveau de zoom 4")
    time.sleep(2)
    voice_commands("zoom level 5", "Niveau de zoom 5")
    time.sleep(2)
    voice_commands("zoom level 1", "Niveau de zoom 1")
    time.sleep(2)
    voice_commands("More options", "Plus D'Options")
    time.sleep(2)
    voice_commands("Aspect ratio", "Format Image")
    time.sleep(2)
    voice_commands("4 by 3", "4 par 3")
    time.sleep(2)
    voice_commands("Aspect ratio", "Format Image")
    time.sleep(2)
    voice_commands("16 by 9", "16 par 9")
    time.sleep(2)
    voice_commands("Field of view", "Champ de vision")
    time.sleep(2)
    voice_commands("Image Resolution", "R??solution de l'image")
    time.sleep(2)
    voice_commands("Low", "Basse")
    time.sleep(2)
    voice_commands("Image Resolution", "R??solution de l'image")
    time.sleep(2)
    voice_commands("High", "Haute")
    time.sleep(2)
    voice_commands("Video Resolution", "R??solution vid??o")
    time.sleep(2)
    voice_commands("Low", "Basse")
    time.sleep(2)
    voice_commands("Video Resolution", "R??solution vid??o")
    time.sleep(2)
    voice_commands("High", "Haute")
    time.sleep(2)
    voice_commands("Frame rate", "Fr??quence d'images")
    time.sleep(2)
    voice_commands("15", "15")
    time.sleep(2)
    voice_commands("Frame rate", "Fr??quence d'images")
    time.sleep(2)
    voice_commands("30", "30")
    time.sleep(2)
    voice_commands("Frame rate", "Fr??quence d'images")
    time.sleep(2)
    voice_commands("25", "25")
    time.sleep(2)
    voice_commands("Video Stabilization", "Stabilisation d'image")
    time.sleep(2)
    voice_commands("Hide Options", "Masquer Les Options")
    time.sleep(2)
    voice_commands("Record Video", "Enregistrer une vid??o")
    time.sleep(10)
    voice_commands("Stop Recording", "Arr??ter l'enregistrement")
    time.sleep(2)
    voice_commands("Preview", "APER??U")
    time.sleep(2)
    voice_commands("Video Pause", "Pause Vid??o")
    time.sleep(2)
    voice_commands("Video Play", "Lecture Vid??o")
    time.sleep(2)
    voice_commands("Video Rewind", "Rembobiner Vid??o")
    time.sleep(2)
    voice_commands("Video Forward", "Avancer Vid??o")
    time.sleep(2)
    voice_commands("Video Stop", "Arr??ter Vid??o")
    time.sleep(2)
    voice_commands("Navigate Home", "Menu Principal")
    time.sleep(2)

    workbook.close()
