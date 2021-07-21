from playsound import playsound
import time, datetime
import os
import subprocess
import xlsxwriter
import re


TC_Name = "MyFiles_Indonesian"
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

def test_MyFiles_Indonesian():

    device_wakeup()
    close_recent_applications()
    time.sleep(2)

    # first level UI
    voice_commands("Navigate Home", "Kembali Ke Awal")
    time.sleep(2)
    voice_commands("My Files", "File Saya")
    time.sleep(2)
    voice_commands("My Media", "Media Saya")
    time.sleep(2)
    voice_commands("Navigate back", "Satu Tingkat Kembali")
    time.sleep(2)
    voice_commands("My Documents", "Dokumen Saya")
    time.sleep(2)
    voice_commands("Navigate back", "Satu Tingkat Kembali")
    time.sleep(2)
    voice_commands("My Photos", "Foto Saya")
    time.sleep(2)
    voice_commands("Navigate back", "Satu Tingkat Kembali")
    time.sleep(2)
    voice_commands("My Downloads", "Unduhan Saya")
    time.sleep(2)
    voice_commands("Navigate back", "Satu Tingkat Kembali")
    time.sleep(2)
    #voice_commands("My Drives", "Drive Saya")
    #time.sleep(2)
    #voice_commands("My device", "Perangkat Saya")
    #time.sleep(2)
    #voice_commands("Navigate back", "Satu Tingkat Kembali")
    #time.sleep(2)
    #voice_commands("Navigate back", "Satu Tingkat Kembali")
    #time.sleep(2)
    #voice_commands("Recent Applications", "Aplikasi Terbaru")
    #time.sleep(2)
    #voice_commands("Dismiss All", "Dismiss all")
    #time.sleep(2)
    voice_commands("Navigate HOME", "Kembali Ke Awal")
    time.sleep(2)
    voice_commands("My CAMERA", "Kameraku Saya")
    time.sleep(2)
    voice_commands("Take Photo", "Memotret")
    time.sleep(10)
    voice_commands("Navigate Home", "Kembali Ke Awal")
    time.sleep(2)
    voice_commands("MY FILES", "File Saya")
    time.sleep(2)
    voice_commands("MY PHOTOS", "Foto Saya")
    time.sleep(2)
    voice_commands("CAMERA", "Camera")
    time.sleep(4)
    voice_commands("SELECT ITEM 1", "Pilih Nomer 1")
    time.sleep(2)
    voice_commands("ZOOM LEVEL 2", "Tingkat zoom 2")
    time.sleep(2)
    voice_commands("ZOOM LEVEL 1", "Tingkat zoom 1")
    time.sleep(2)
    voice_commands("ZOOM LEVEL 4", "Tingkat zoom 4")
    time.sleep(2)
    voice_commands("ZOOM LEVEL 5", "Tingkat zoom 5")
    time.sleep(2)
    voice_commands("ZOOM LEVEL 3", "Tingkat zoom 3")
    time.sleep(2)
    voice_commands("FREEZE WINDOW", "Bekukan Layar")
    time.sleep(2)
    voice_commands("CONTROL WINDOW", "Aktifkan Layar")
    time.sleep(2)
    voice_commands("RESET WINDOW", "ulang jendela")
    time.sleep(2)
    voice_commands("Navigate back", "Satu Tingkat Kembali")
    time.sleep(2)
    voice_commands("Navigate back", "Satu Tingkat Kembali")
    time.sleep(2)
    voice_commands("MORE OPTIONS", "Pilihan Lain")
    time.sleep(2)
    voice_commands("SORT FILES", "Pilah File")
    time.sleep(2)
    voice_commands("BY DATE", "Dengan Tanggal")
    time.sleep(2)
    voice_commands("Sort Files", "Pilah File")
    time.sleep(2)
    voice_commands("By Name", "Dengan nama")
    time.sleep(2)
    voice_commands("Hide Options", "Sembunyikan Pilihan")
    time.sleep(2)
    voice_commands("Camera", "Camera")
    time.sleep(2)
    voice_commands("MORE OPTIONS", "Pilihan Lain")
    time.sleep(2)
    voice_commands("EDIT MODE", "Moda Edit")
    time.sleep(3)
    voice_commands("SELECT ITEM 1", "Pilih Nomer 1")
    time.sleep(3)
    voice_commands("DELETE SELECTED", "Hapus Gambar Pilihan")
    time.sleep(2)
    voice_commands("CONFIRM DELETION", "KONFIRMASI HAPUS")
    time.sleep(2)

    #voice_commands("Recent Applications", "Aplikasi Terbaru")
    #time.sleep(2)
    #voice_commands("Dismiss All", "Dismiss all")
    #time.sleep(2)
    #voice_commands("MY FILES", "File Saya")
    #time.sleep(2)
    voice_commands("Navigate back", "Satu Tingkat Kembali")
    time.sleep(2)
    voice_commands("Navigate back", "Satu Tingkat Kembali")
    time.sleep(2)

    voice_commands("MY DOCUMENTS", "Dokumen Saya")
    time.sleep(3)

    voice_commands("SELECT ITEM 1", "Pilih Nomer 1")
    time.sleep(2)
    voice_commands("ZOOM LEVEL 3", "Tingkat zoom 3")
    time.sleep(2)
    #voice_commands("ZOOM LEVEL 1", "Tingkat zoom 1")
    #time.sleep(2)
    voice_commands("ZOOM LEVEL 4", "Tingkat zoom 4")
    time.sleep(2)
    voice_commands("ZOOM LEVEL 5", "Tingkat zoom 5")
    time.sleep(2)
    voice_commands("ZOOM LEVEL 2", "Tingkat zoom 2")
    time.sleep(2)
    voice_commands("FREEZE WINDOW", "Bekukan Layar")
    time.sleep(2)
    voice_commands("CONTROL WINDOW", "Aktifkan Layar")
    time.sleep(2)
    voice_commands("Reset window", "Ulang Jendela")
    time.sleep(2)
    voice_commands("GO TO PAGE 10", "Pergi ke halaman 10")
    time.sleep(2)

    workbook.close()
