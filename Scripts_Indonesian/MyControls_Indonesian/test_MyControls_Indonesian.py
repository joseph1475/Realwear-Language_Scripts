from playsound import playsound
import time, datetime
import os
import subprocess
import xlsxwriter
import re


TC_Name = "MyControls_Indonesian"
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

def test_MyControls_Indonesian():
    device_wakeup()
    close_recent_applications()
    time.sleep(2)
    voice_commands("Navigate Home","Kembali Ke Awal")
    time.sleep(2)
    voice_commands("My Controls","Pengaturan Pemakai")
    time.sleep(2)
    voice_commands("Mouse","Tetikuis")
    time.sleep(2)
    voice_commands("Auto Rotate","Rotasi Otomatis")
    time.sleep(2)
    voice_commands("Action Button","Tombol Tindakan")
    #time.sleep(1)
    #voice_commands("Noise Capture","tangkap suara latar")
    #time.sleep(2)
    #voice_commands("Action Button","Tombol Tindakan")
    #time.sleep(1)
    #voice_commands("Homescreen","Tampilan Depan")
    time.sleep(3)
    voice_commands("Power Options","Pilihan Power")
    time.sleep(4)
    voice_commands("My Controls","Pengaturan Pemakai")
    time.sleep(2)
    voice_commands("WearML Indicators","Indikator WearML")
    time.sleep(1)
    voice_commands("Fade 5 Seconds","Memudar 5 Detik")
    time.sleep(2)
    voice_commands("WearML Indicators","Indikator WearML")
    time.sleep(1)
    voice_commands("Fade 10 Seconds","Memudar 10 Detik")
    time.sleep(2)
    #voice_commands("WearML Indicators","Indikator WearML")
    #time.sleep(1)
    #voice_commands("Fade never","Tidak Memudar")
    #time.sleep(2)
    voice_commands("WearML Indicators","Indikator WearML")
    time.sleep(1)
    voice_commands("Fade 3 Seconds","Memudar 3 Detik")
    time.sleep(4)
    voice_commands("Navigate Home","Kembali Ke Awal")
    time.sleep(2)
    voice_commands("My Controls","Pengaturan Pemakai")
    time.sleep(2)
    voice_commands("Bluetooth","Bluetooth")
    time.sleep(1)
    voice_commands("Bluetooth settings","Pengaturan Bluetooth")
    time.sleep(3)
    voice_commands("Navigate Home","Kembali Ke Awal")
    time.sleep(2)
    voice_commands("My Controls","Pengaturan Pemakai")
    time.sleep(3)

    voice_commands("Wireless Network","Jaringan nirkabel")
    #time.sleep(1)
    #voice_commands("Wireless Network settings","Pengaturan Jaringan Nirkabel")
    time.sleep(5)
    #voice_commands("Navigate Back","Satu Tingkat Kembali")
    #time.sleep(2)
    #voice_commands("Navigate Home","Kembali Ke Awal")
    #time.sleep(2)
    #voice_commands("My Controls","Pengaturan Pemakai")
    #time.sleep(2)
    voice_commands("Preferred Network","Jaringan Pilihan")
    time.sleep(2)
    voice_commands("Navigate Back","Satu Tingkat Kembali")
    time.sleep(3)
    voice_commands("Navigate Home","Kembali Ke Awal")
    time.sleep(2)
    voice_commands("My Controls","Pengaturan Pemakai")
    time.sleep(5)
    #voice_commands("Flashlight","Senter")
    #time.sleep(5)
    #voice_commands("Flashlight","Senter")
    #time.sleep(2)
    #voice_commands("Navigate Back","Satu Tingkat Kembali")
    ##time.sleep(2)
    #voice_commands("My Controls","Pengaturan Pemakai")
    time.sleep(2)
    voice_commands("Brightness","Kecerahan")
    time.sleep(0)
    voice_commands("Set level 1","atur tingkat 1")
    time.sleep(5)
    voice_commands("Brightness","Kecerahan")
    time.sleep(0)
    voice_commands("Set level 2","atur tingkat 2")
    time.sleep(5)
    voice_commands("Brightness","Kecerahan")
    time.sleep(0)
    voice_commands("Set level 3","atur tingkat 3")
    time.sleep(5)
    voice_commands("Brightness","Kecerahan")
    time.sleep(0)
    voice_commands("Set level 5","atur tingkat 5")
    time.sleep(5)
    voice_commands("Brightness","Kecerahan")
    time.sleep(0)
    voice_commands("Set level 6","atur tingkat 6")
    time.sleep(5)
    voice_commands("Brightness","Kecerahan")
    time.sleep(0)
    voice_commands("Set level 7","atur tingkat 7")
    time.sleep(5)
    voice_commands("Brightness","Kecerahan")
    time.sleep(0)
    voice_commands("Set level 8","atur tingkat 8")
    time.sleep(5)
    voice_commands("Brightness","Kecerahan")
    time.sleep(0)
    voice_commands("Set level 9","atur tingkat 9")
    time.sleep(5)
    voice_commands("Brightness","Kecerahan")
    time.sleep(0)
    voice_commands("Set level 10","atur tingkat 10")
    time.sleep(5)
    voice_commands("Brightness","Kecerahan")
    time.sleep(0)
    voice_commands("Set level 4","atur tingkat 4")
    time.sleep(5)
    voice_commands("Navigate Home","Kembali Ke Awal")
    time.sleep(2)
    voice_commands("My Controls","Pengaturan Pemakai")
    time.sleep(2)
    voice_commands("Volume","Kekerasan Suara")
    time.sleep(0)
    voice_commands("Set level 1","atur tingkat 1")
    time.sleep(5)
    voice_commands("Volume","Kekerasan Suara")
    time.sleep(0)
    voice_commands("Set level 2","atur tingkat 2")
    time.sleep(5)
    voice_commands("Volume","Kekerasan Suara")
    time.sleep(0)
    voice_commands("Set level 3","atur tingkat 3")
    time.sleep(5)
    voice_commands("Volume","Kekerasan Suara")
    time.sleep(0)
    voice_commands("Set level 4","atur tingkat 4")
    time.sleep(5)
    voice_commands("Volume","Kekerasan Suara")
    time.sleep(0)
    voice_commands("Set level 6","atur tingkat 6")
    time.sleep(5)
    voice_commands("Volume","Kekerasan Suara")
    time.sleep(0)
    voice_commands("Set level 7","atur tingkat 7")
    time.sleep(5)
    voice_commands("Volume","Kekerasan Suara")
    time.sleep(0)
    voice_commands("Set level 8","atur tingkat 8")
    time.sleep(5)
    voice_commands("Volume","Kekerasan Suara")
    time.sleep(0)
    voice_commands("Set level 9","atur tingkat 9")
    time.sleep(5)
    voice_commands("Volume","Kekerasan Suara")
    time.sleep(0)
    voice_commands("Set level 10","atur tingkat 10")
    time.sleep(5)
    voice_commands("Volume","Kekerasan Suara")
    time.sleep(0)
    voice_commands("Set level 5","atur tingkat 5")
    time.sleep(5)
    voice_commands("Color mode","Modus warna")
    time.sleep(2)
    voice_commands("Color mode","Modus warna")
    time.sleep(2)
    voice_commands("Dictation","Dikte")
    time.sleep(2)
    voice_commands("Dictation","Dikte")
    time.sleep(2)
    voice_commands("Help command","Ikon Bantuan")
    time.sleep(2)
    voice_commands("Help command","Ikon Bantuan")
    time.sleep(2)
    voice_commands("Layout mode","Moda Tata Letak")
    time.sleep(2)
    voice_commands("Page down","HALAMAN KE BAWAH")
    time.sleep(2)
    voice_commands("Layout mode","Moda Tata Letak")
    time.sleep(2)
    voice_commands("More settings","Pengaturan Lain")
    time.sleep(2)
    voice_commands("Navigate back","Satu Tingkat Kembali")
    time.sleep(2)

    workbook.close()



