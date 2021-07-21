from playsound import playsound
import time, datetime
import os
import subprocess
import xlsxwriter
import re


TC_Name = "MyFiles_French"
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

def test_MyFiles_French():
    device_wakeup()
    close_recent_applications()
    time.sleep(2)

    # first level UI
    voice_commands("Navigate Home", "Menu Principal")
    time.sleep(2)
    voice_commands("My Files", "Mes Fichiers")
    time.sleep(2)
    voice_commands("My Media", "Mes Médias")
    time.sleep(2)
    voice_commands("Navigate back", "Retour En Arrière")
    time.sleep(2)
    voice_commands("My Documents", "Mes Documents")
    time.sleep(2)
    voice_commands("Navigate back", "Retour En Arrière")
    time.sleep(2)
    voice_commands("My Photos", "Mes Photos")
    time.sleep(2)
    voice_commands("Navigate back", "Retour En Arrière")
    time.sleep(2)
    voice_commands("My Downloads", "Mes Téléchargements")
    time.sleep(2)
    voice_commands("Navigate back", "Retour En Arrière")
    time.sleep(2)
    voice_commands("My Drives", "Mes Disques")
    time.sleep(2)
    voice_commands("My device", "Mon Appareil")
    time.sleep(2)
    voice_commands("Back One level", "Retour d'un niveau")
    time.sleep(2)
    voice_commands("Navigate back", "Retour En Arrière")
    time.sleep(2)
    voice_commands("Recent Applications", "Applications récentes")
    time.sleep(3)
    voice_commands("Dismiss All", "Tout rejeter")
    time.sleep(2)
    voice_commands("Navigate HOME", "Menu Principal")
    time.sleep(2)
    voice_commands("My CAMERA", "Mon appareil photo")
    time.sleep(2)
    voice_commands("Take Photo", "Prendre une Photo")
    time.sleep(8)
    voice_commands("Navigate Home", "Menu Principal")
    time.sleep(2)
    voice_commands("MY FILES", "Mes Fichiers")
    time.sleep(2)
    voice_commands("MY PHOTOS", "Mes Photos")
    time.sleep(2)
    voice_commands("CAMERA", "Camera")
    time.sleep(4)
    voice_commands("SELECT ITEM 1", "Sélectionnez l'élément 1")
    time.sleep(2)
    voice_commands("ZOOM LEVEL 2", "Niveau de zoom 2")
    time.sleep(2)
    voice_commands("ZOOM LEVEL 1", "Niveau de zoom 1")
    time.sleep(2)
    voice_commands("ZOOM LEVEL 4", "Niveau de zoom 4")
    time.sleep(2)
    voice_commands("ZOOM LEVEL 5", "Niveau de zoom 5")
    time.sleep(2)
    voice_commands("ZOOM LEVEL 3", "Niveau de zoom 3")
    time.sleep(2)
    voice_commands("FREEZE WINDOW", "Gel écran")
    time.sleep(2)
    voice_commands("CONTROL WINDOW", "Déblocage écran")
    time.sleep(2)
    voice_commands("RESET WINDOW", "Réinitialiser la fenêtre")
    time.sleep(2)
    voice_commands("Navigate back", "Retour En Arrière")
    time.sleep(2)
    voice_commands("Back one level", "Retour d'un niveau")
    time.sleep(2)
    voice_commands("MORE OPTIONS", "Plus D'Options")
    time.sleep(2)
    voice_commands("SORT FILES", "Trier les fichiers")
    time.sleep(2)
    voice_commands("BY DATE", "Par date")
    time.sleep(2)
    voice_commands("Sort Files", "Trier les fichiers")
    time.sleep(2)
    voice_commands("By Name", "Par nom")
    time.sleep(2)
    voice_commands("Hide Options", "Masquer Les Options")
    time.sleep(2)
    voice_commands("Camera", "Camera")
    time.sleep(2)
    voice_commands("MORE OPTIONS", "Plus D'Options")
    time.sleep(2)
    voice_commands("EDIT MODE", "Mode édition")
    time.sleep(3)
    voice_commands("SELECT ITEM 1", "Sélectionnez l'élément 1")
    time.sleep(3)
    voice_commands("DELETE SELECTED", "Supprimer la sélection")
    time.sleep(2)
    voice_commands("CONFIRM DELETION", "CONFIRMER")
    time.sleep(2)
    voice_commands("Recent Applications", "Applications récentes")
    time.sleep(2)
    voice_commands("Dismiss All", "Tout rejeter")
    time.sleep(2)
    voice_commands("MY FILES", "Mes Fichiers")
    time.sleep(2)
    voice_commands("MY DOCUMENTS", "Mes Documents")
    time.sleep(3)
    voice_commands("SELECT ITEM 1", "Sélectionnez l'élément 1")
    time.sleep(5)
    voice_commands("ZOOM LEVEL 3", "Niveau de zoom 3")
    time.sleep(2)
    voice_commands("ZOOM LEVEL 1", "Niveau de zoom 1")
    time.sleep(2)
    voice_commands("ZOOM LEVEL 4", "Niveau de zoom 4")
    time.sleep(2)
    voice_commands("ZOOM LEVEL 5", "Niveau de zoom 5")
    time.sleep(2)
    voice_commands("ZOOM LEVEL 2", "Niveau de zoom 2")
    time.sleep(2)
    voice_commands("FREEZE WINDOW", "Gel écran")
    time.sleep(2)
    voice_commands("CONTROL WINDOW", "Déblocage écran")
    time.sleep(2)
    voice_commands("Reset window", "Réinitialiser la fenêtre")
    time.sleep(2)
    #voice_commands("GO TO PAGE 10", "aller à la page 10")
    #time.sleep(2)
    workbook.close()
