from playsound import playsound
import time, datetime
import os
import subprocess
import xlsxwriter
import re


TC_Name = "GeneralSettings_Indonesian"
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

def test_General_Setings_Indonesian():

    device_wakeup()
    close_recent_applications()
    time.sleep(2)

    # Network and Internet options testing
    voice_commands("Navigate Home","Kembali Ke Awal")
    time.sleep(2)
    voice_commands("My Programs","Program Saya")
    time.sleep(2)
    voice_commands("Settings","Setelan")
    time.sleep(4)
    voice_commands("Network and Internet","Jaringan dan internet")
    time.sleep(2)
    #voice_commands("Wireless Networks","")
    #time.sleep(2)
    #voice_commands("Navigate back",""Satu Tingkat Kembali")
    #time.sleep(3)
    voice_commands("Navigate back","Satu Tingkat Kembali")
    time.sleep(2)
    voice_commands("Connected devices","Perangkat tersambung")
    time.sleep(2)
    #voice_commands("USB","USB")
    #time.sleep(2)
    #voice_commands("Navigate Back","Satu Tingkat Kembali")
    #time.sleep(2)
    voice_commands("Pair new device","Sambungkan perangkat baru")
    time.sleep(2)
    voice_commands("Navigate Back","Satu Tingkat Kembali")
    time.sleep(2)
    voice_commands("Connection preferences","Preferensi sambungan")
    time.sleep(2)
    voice_commands("Navigate Back","Satu Tingkat Kembali")
    time.sleep(2)
    voice_commands("Navigate Back","Satu Tingkat Kembali")
    time.sleep(2)
    voice_commands("Apps and notifications","Aplikasi dan notifikasi")
    time.sleep(2)
    voice_commands("Navigate Back","Satu Tingkat Kembali")
    time.sleep(2)
    #voice_commands("battery","")
    #time.sleep(2)
    #voice_commands("Navigate Back","")
    #time.sleep(2)
    voice_commands("Display","Tampilan")
    time.sleep(2)
    voice_commands("Brightness level","Tingkat kecerahan")
    time.sleep(2)
    voice_commands("Navigate Back","Satu Tingkat Kembali")
    time.sleep(2)
    voice_commands("Night Light","Cahaya Malam")
    time.sleep(2)
    voice_commands("Navigate back","Satu Tingkat Kembali")
    time.sleep(2)

    #voice_commands("Screen timeout","Waktu tunggu layar")
    #time.sleep(2)
    #voice_commands("30 seconds","30 detik")
    #time.sleep(2)
    #voice_commands("Screen timeout","")
    #time.sleep(2)
    #voice_commands("5 minutes","")
    #time.sleep(2)
    #voice_commands("Screen timeout","")
    #time.sleep(2)
    #voice_commands("10 minutes","")
    #time.sleep(2)
    #voice_commands("Screen timeout","")
    #time.sleep(2)
    #voice_commands("30 minutes","")
    #time.sleep(2)
    #voice_commands("Screen timeout","")
    #time.sleep(2)
    #voice_commands("CANCEL","")
    #time.sleep(2)

    voice_commands("Page Down","Halaman Ke Bawah")
    time.sleep(2)
    voice_commands("Advanced","Lanjutan")
    time.sleep(2)
    voice_commands("Font size","Ukuran font")
    time.sleep(2)
    voice_commands("Navigate Back","Satu Tingkat Kembali")
    time.sleep(2)
    voice_commands("Navigate Back","Satu Tingkat Kembali")
    time.sleep(2)
    voice_commands("Sound","Suara")
    time.sleep(2)
    voice_commands("Navigate Back","Satu Tingkat Kembali")
    time.sleep(2)
    voice_commands("Page down","Halaman Ke Bawah")
    time.sleep(2)
    voice_commands("Storage","Penyimpanan")
    time.sleep(2)
    voice_commands("Navigate Back","Satu Tingkat Kembali")
    time.sleep(2)
    voice_commands("Privacy","Privasi")
    time.sleep(2)
    voice_commands("Accessibility usage","Penggunaan aksesibilitas")
    time.sleep(2)
    voice_commands("ok","YA")
    time.sleep(2)
    voice_commands("Permission manager","Pengelola izin")
    time.sleep(2)
    voice_commands("navigate back","Satu Tingkat Kembali")
    time.sleep(2)
    voice_commands("Show passwords","Tampilkan sandi")
    time.sleep(2)
    voice_commands("Show passwords","Tampilkan sandi")
    time.sleep(2)
    voice_commands("Lock screen","Layar kunci")
    time.sleep(2)
    voice_commands("Navigate Back","Satu Tingkat Kembali")
    time.sleep(2)
    voice_commands("Navigate Back","Satu Tingkat Kembali")
    time.sleep(2)
    voice_commands("Location","Lokasi")
    time.sleep(2)
    voice_commands("Use Location","Gunakan lokasi")
    time.sleep(2)
    voice_commands("Use Location","Gunakan lokasi")
    time.sleep(2)
    voice_commands("Navigate back","Satu Tingkat Kembali")
    time.sleep(2)
    voice_commands("Security","Keamanan")
    time.sleep(2)
    voice_commands("Navigate back","Satu Tingkat Kembali")
    time.sleep(2)
    voice_commands("Accounts","Akun")
    time.sleep(2)
    voice_commands("Navigate Back","Satu Tingkat Kembali")
    time.sleep(2)
    voice_commands("Accessibility","Aksesibilitas")
    time.sleep(2)
    voice_commands("Navigate back","Satu Tingkat Kembali")
    time.sleep(2)
    voice_commands("Realwear","RealWear")
    time.sleep(2)
    voice_commands("Bluetooth Mode","Moda Bluetooth")
    time.sleep(2)
    voice_commands("Bluetooth Mode","Moda Bluetooth")
    time.sleep(2)

    voice_commands("Storage Mode","Moda Penyimpanan")
    time.sleep(2)
    voice_commands("Storage Mode","Moda Penyimpanan")
    time.sleep(2)
    voice_commands("CameraLow light","Kamera Cahaya Rendah")
    time.sleep(2)
    voice_commands("CameraLow light","Kamera Cahaya Rendah")
    time.sleep(2)
    voice_commands("CameraField of View","Kamera Bidang Tampilan")
    time.sleep(2)
    voice_commands("CameraField of View","Kamera Bidang Tampilan")
    time.sleep(2)
    voice_commands("CameraImage Stabilization","Kamera Stabilisasi Gambar")
    time.sleep(2)
    voice_commands("CameraImage Stabilization","Kamera Stabilisasi Gambar")
    time.sleep(2)
    voice_commands("Navigate Back","Satu Tingkat Kembali")
    time.sleep(2)
    voice_commands("Page down","Halaman Ke Bawah")
    time.sleep(2)
    voice_commands("System","Sistem")
    time.sleep(2)
    voice_commands("Navigate back","Satu Tingkat Kembali")
    time.sleep(2)
    """
    voice_commands("About Device","Tentang ponsel")
    time.sleep(2)
    voice_commands("Device name","nama perangkat")
    time.sleep(2)
    voice_commands("navigate Back","Satu Tingkat Kembali")
    time.sleep(2)
    voice_commands("Oke","oke")
    time.sleep(2)
    voice_commands("Cancel","batal")
    time.sleep(2)
    voice_commands("Emergency information","informasi darurat")
    time.sleep(2)
    voice_commands("navigate Back","Satu Tingkat Kembali")
    time.sleep(2)
    voice_commands("Model","Model")
    time.sleep(2)
    voice_commands("navigate Back","Satu Tingkat Kembali")
    time.sleep(2)
    voice_commands("Page Down","Halaman Ke Bawah")
    time.sleep(2)
    voice_commands("Android version","Versi Android")
    time.sleep(2)
    voice_commands("Navigate back","Satu Tingkat Kembali")
    time.sleep(2)
    voice_commands("Page Down","Halaman Ke Bawah")
    time.sleep(2)
    voice_commands("Build number","Nomor versi")
    time.sleep(2)
    voice_commands("Navigate back","Satu Tingkat Kembali")
    time.sleep(2)
    """
    workbook.close()