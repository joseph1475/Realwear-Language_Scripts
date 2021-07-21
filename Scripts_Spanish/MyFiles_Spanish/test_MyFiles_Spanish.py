from playsound import playsound
import time, datetime
import os
import subprocess
import xlsxwriter
import re


TC_Name = "MyFiles_Spanish"
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

def test_MyFiles_Spanish():
    device_wakeup()
    close_recent_applications()
    time.sleep(2)

    #first level UI
    voice_commands("Navigate Home","Regresar a inicio")
    time.sleep(2)
    voice_commands("My Files","Mis Archivos")
    time.sleep(2)
    voice_commands("My Media","Mi Informacion")
    time.sleep(2)
    voice_commands("Navigate back","Ir atrás")
    time.sleep(2)
    voice_commands("My Documents","Mis Documentos")
    time.sleep(2)
    voice_commands("Navigate back","Ir atrás")
    time.sleep(2)
    voice_commands("My Photos","Mis Fotos")
    time.sleep(2)
    voice_commands("Navigate back","Ir atrás")
    time.sleep(2)
    voice_commands("My Downloads","Mis Descargas")
    time.sleep(2)
    voice_commands("Navigate back","Ir atrás")
    time.sleep(2)
    voice_commands("My Drives","Mis Unidades")
    time.sleep(2)
    #voice_commands("My device","Mi Dispositivo")
    #time.sleep(2)
    #voice_commands("Back One level","Volver un nivel")
    #time.sleep(2)
    voice_commands("Navigate back","Ir atrás")
    time.sleep(2)
    voice_commands("Recent Applications","Aplicaciones recientes")
    time.sleep(3)
    voice_commands("Dismiss All","Descartar todo")
    time.sleep(2)
    voice_commands("Navigate HOME","Regresar a inicio")
    time.sleep(2)
    voice_commands("My CAMERA","Mi Cámara")
    time.sleep(2)
    voice_commands("Take Photo","Tomar foto")
    time.sleep(8)
    voice_commands("Navigate Home","Regresar a inicio")
    time.sleep(2)
    voice_commands("MY FILES","Mis Archivos")
    time.sleep(2)
    voice_commands("MY PHOTOS","Mis Fotos")
    time.sleep(2)
    voice_commands("CAMERA","Camera")
    time.sleep(4)
    voice_commands("SELECT ITEM 1","Elemento 1")
    time.sleep(2)
    voice_commands("ZOOM LEVEL 2","Nivel de zoom 2")
    time.sleep(2)
    voice_commands("ZOOM LEVEL 1","Nivel de zoom 1")
    time.sleep(2)
    voice_commands("ZOOM LEVEL 4","Nivel de zoom 4")
    time.sleep(2)
    voice_commands("ZOOM LEVEL 5","Nivel de zoom 5")
    time.sleep(2)
    voice_commands("ZOOM LEVEL 3","Nivel de zoom 3")
    time.sleep(2)
    voice_commands("FREEZE WINDOW","Congelar ventana")
    time.sleep(2)
    voice_commands("CONTROL WINDOW","Controlar ventana")
    time.sleep(2)
    voice_commands("RESET WINDOW","Reiniciar Ventana")
    time.sleep(2)
    voice_commands("Navigate back","Ir atrás")
    time.sleep(2)
    voice_commands("Back one level","Volver un nivel")
    time.sleep(2)
    voice_commands("MORE OPTIONS","Mas Opciones")
    time.sleep(2)
    voice_commands("SORT FILES","Ordenar archivos")
    time.sleep(2)
    voice_commands("BY DATE","Por fecha")
    time.sleep(2)
    voice_commands("Sort Files","Ordenar archivos")
    time.sleep(2)
    voice_commands("By Name","Por nombre")
    time.sleep(2)
    voice_commands("Hide Options","Ocultar Opciones")
    time.sleep(2)
    voice_commands("Camera","Camera")
    time.sleep(2)
    voice_commands("MORE OPTIONS","Mas Opciones")
    time.sleep(2)
    voice_commands("EDIT MODE","Editar")
    time.sleep(3)
    voice_commands("SELECT ITEM 1","Elemento 1")
    time.sleep(3)
    voice_commands("DELETE SELECTED","Eliminar selección")
    time.sleep(2)
    voice_commands("CONFIRM DELETION","CONFIRMAR")
    time.sleep(2)
    voice_commands("Recent Applications","Aplicaciones recientes")
    time.sleep(2)
    voice_commands("Dismiss All","Descartar todo")
    time.sleep(2)
    voice_commands("MY FILES","Mis Archivos")
    time.sleep(2)
    voice_commands("MY DOCUMENTS","Mis Documentos")
    time.sleep(3)
    voice_commands("SELECT ITEM 1","Elemento 1")
    time.sleep(5)
    voice_commands("ZOOM LEVEL 3","Nivel de zoom 3")
    time.sleep(2)
    voice_commands("ZOOM LEVEL 1","Nivel de zoom 1")
    time.sleep(2)
    voice_commands("ZOOM LEVEL 4","Nivel de zoom 4")
    time.sleep(2)
    voice_commands("ZOOM LEVEL 5","Nivel de zoom 5")
    time.sleep(2)
    voice_commands("ZOOM LEVEL 2","Nivel de zoom 2")
    time.sleep(2)
    voice_commands("FREEZE WINDOW","Congelar ventana")
    time.sleep(2)
    voice_commands("CONTROL WINDOW","Controlar ventana")
    time.sleep(2)
    voice_commands("Reset window","Reiniciar Ventana")
    time.sleep(2)
    voice_commands("GO TO PAGE 10","Ir a página 10")
    time.sleep(2)

    workbook.close()
