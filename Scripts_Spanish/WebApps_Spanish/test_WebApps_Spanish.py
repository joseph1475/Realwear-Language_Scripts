from playsound import playsound
import time, datetime
import os
import subprocess
import xlsxwriter
import re


TC_Name = "WebApps_Spanish"
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


def Accept_cookies():
    try:
        playsound('Tones/Accept.wav')
        time.sleep(4)
    except:
        playsound('Tones/Accept.mp3')
        time.sleep(4)

def device_wakeup():
    adb_wakeup = "adb shell input keyevent KEYCODE_WAKEUP"
    os.system(adb_wakeup)

def test_WebApps_Spanish():

    device_wakeup()
    close_recent_applications()
    time.sleep(2)

    #launching web apps
    voice_commands("Navigate Home","Regresar a inicio")
    time.sleep(2)
    voice_commands("My Programs","Mis programas")
    time.sleep(2)
    voice_commands("Web Apps","Aplicaciones Web")
    time.sleep(2)
    #youtube verifications
    voice_commands("Youtube","Youtube")
    time.sleep(2)
    #voice_commands("Mouse Click")
    #time.sleep(4)
    voice_commands("Refresh Page","Actualizar p??gina")
    time.sleep(4)
    voice_commands("More Options","Mas Opciones")
    time.sleep(2)
    voice_commands("Mouse","Rat??n")
    time.sleep(2)
    #voice_commands("Hide Options")
    #time.sleep(3)
    #voice_commands("Browser Back")
    #time.sleep(3)
    #voice_commands("Browser Forward")
    #time.sleep(3)
    #voice_commands("More Options")
    #time.sleep(2)
    voice_commands("Scan code","Escanear Codigo")
    time.sleep(2)
    voice_commands("Navigate Back","Ir atr??s")
    time.sleep(4)
    voice_commands("Add to Bookmarks","A??adir a marcadores")
    time.sleep(2)
    voice_commands("Cancel","CANCELAR")
    time.sleep(2)
    voice_commands("Add to Bookmarks","A??adir a marcadores")
    time.sleep(2)
    voice_commands("Save Bookmark","GUARDAR MARCADOR")
    time.sleep(2)
    voice_commands("Hide Options","Ocultar Opciones")
    time.sleep(2)
    """
    voice_commands("Recent Applications")
    time.sleep(2)
    voice_commands("Dismiss All")
    time.sleep(2)
    
    #keyboard launch
    voice_commands("My Programs")
    time.sleep(2)
    voice_commands("Web Apps")
    time.sleep(2)
    voice_commands("Youtube")
    time.sleep(4)
    voice_commands("More Options")
    time.sleep(2)
    voice_commands("Mouse")
    time.sleep(2)
    voice_commands("hide Options")
    time.sleep(2)
    voice_commands("Select item 2")
    time.sleep(3)
    voice_commands("Symbols")
    time.sleep(3)
    voice_commands("Letters")
    time.sleep(3)
    voice_commands("ABC MODE")
    time.sleep(3)
    voice_commands("Numbers")
    time.sleep(2)
    voice_commands("Navigate Back")
    time.sleep(4)
    voice_commands("Recent Applications")
    time.sleep(2)
    voice_commands("Dismiss All")
    time.sleep(2)
    
    #Realwear
    voice_commands("Navigate Home")
    time.sleep(2)
    voice_commands("My Programs")
    time.sleep(2)
    voice_commands("Web Apps")
    time.sleep(3)
    """
    voice_commands("Navigate back","Ir atr??s")
    time.sleep(5)
    voice_commands("Realwear","Realwear")
    #time.sleep(2)
    #voice_commands("Mouse Click")
    time.sleep(4)
    voice_commands("Refresh Page","Actualizar p??gina")
    time.sleep(2)
    Accept_cookies()
    voice_commands("More Options","Mas Opciones")
    time.sleep(2)
    voice_commands("Mouse","Rat??n")
    time.sleep(2)
    voice_commands("Hide Options","Ocultar Opciones")
    time.sleep(4)
    voice_commands("Select item 3","Elemento 3")
    time.sleep(4)
    #voice_commands("Browser Back")
    #time.sleep(4)
    #voice_commands("Browser Forward")
    #time.sleep(2)
    voice_commands("More Options","Mas Opciones")
    time.sleep(2)
    voice_commands("Scan code","Escanear Codigo")
    time.sleep(2)
    voice_commands("Navigate Back","Ir atr??s")
    time.sleep(4)
    voice_commands("Add to Bookmarks","A??adir a marcadores")
    time.sleep(2)
    voice_commands("Cancel","CANCELAR")
    time.sleep(2)
    voice_commands("Add to Bookmarks","A??adir a marcadores")
    time.sleep(2)
    voice_commands("Save Bookmark","GUARDAR MARCADOR")
    time.sleep(2)
    voice_commands("Hide Options","Ocultar Opciones")
    time.sleep(2)
    """
    voice_commands("Recent Applications")
    time.sleep(2)
    voice_commands("Dismiss All")
    time.sleep(2)

    #Google
    voice_commands("My Programs")
    time.sleep(2)
    voice_commands("Web Apps")
    time.sleep(2)
    """
    voice_commands("Navigate Back","Ir atr??s")
    time.sleep(4)
    voice_commands("Google","Google")
    time.sleep(2)
    voice_commands("Refresh Page","Actualizar p??gina")
    time.sleep(4)
    voice_commands("More Options","Mas Opciones")
    time.sleep(4)
    #voice_commands("Mouse Click")
    #time.sleep(4)
    voice_commands("Mouse","Rat??n")
    time.sleep(2)
    voice_commands("Scan code","Escanear Codigo")
    time.sleep(2)
    voice_commands("Navigate Back","Ir atr??s")
    time.sleep(4)
    voice_commands("Add to Bookmarks","A??adir a marcadores")
    time.sleep(2)
    voice_commands("Cancel","CANCELAR")
    time.sleep(2)
    voice_commands("Add to Bookmarks","A??adir a marcadores")
    time.sleep(2)
    voice_commands("Save Bookmark","GUARDAR MARCADOR")
    time.sleep(4)
    voice_commands("Hide Options","Ocultar Opciones")
    time.sleep(2)
    #voice_commands("Select item 11")
    #time.sleep(2)
    #voice_commands("Select item 12")
    #time.sleep(2)
    #voice_commands("Browser Forward")
    #time.sleep(2)
    #voice_commands("Browser Back")
    #time.sleep(2)
    voice_commands("Recent Applications","Aplicaciones recientes")
    time.sleep(2)
    voice_commands("Dismiss All","Descartar todo")
    time.sleep(2)
    """
    #keyboard launch
    voice_commands("My Programs")
    time.sleep(2)
    voice_commands("Web Apps")
    time.sleep(2)
    voice_commands("Google")
    time.sleep(2)
    voice_commands("More Options")
    time.sleep(2)
    voice_commands("Mouse")
    time.sleep(2)
    voice_commands("hide Options")
    time.sleep(4)
    voice_commands("Select item 6")
    time.sleep(2)
    voice_commands("Symbols")
    time.sleep(2)
    voice_commands("Letters")
    time.sleep(2)
    voice_commands("ABC MODE")
    time.sleep(2)
    voice_commands("Numbers")
    time.sleep(2)
    voice_commands("Navigate Home")
    time.sleep(2)
    """
    workbook.close()
