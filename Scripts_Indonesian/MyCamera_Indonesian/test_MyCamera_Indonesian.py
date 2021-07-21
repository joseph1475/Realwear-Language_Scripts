from playsound import playsound
import time, datetime
import os
import subprocess
import xlsxwriter
import re


TC_Name = "MyCamera_Indonesian"
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

def test_MyCamera_Indonesian():
    device_wakeup()
    close_recent_applications()
    time.sleep(2)
    voice_commands("Navigate Home", "Kembali Ke Awal")
    time.sleep(2)
    voice_commands("My Programs","Program Saya")
    time.sleep(2)
    voice_commands("My Camera","Kameraku Saya")
    time.sleep(2)
    """
    voice_commands("Take Photo", "Memotret")
    time.sleep(3)
    voice_commands("Preview", "PRATINJAU")
    time.sleep(2)
    voice_commands("zoom level 2", "Tingkat zoom 2")
    time.sleep(2)
    voice_commands("zoom level 3", "Tingkat zoom 3")
    time.sleep(2)
    voice_commands("zoom level 4", "Tingkat zoom 4")
    time.sleep(2)
    voice_commands("zoom level 5", "Tingkat zoom 5")
    time.sleep(2)
    voice_commands("Freeze Window", "Bekukan Layar")
    time.sleep(2)
    voice_commands("Control Window", "aktifkan layar")
    time.sleep(2)
    voice_commands("zoom level 1", "Tingkat zoom 1")
    time.sleep(2)
    voice_commands("Delete", "Hapus")
    time.sleep(2)
    voice_commands("Confirm", "KONFIRMASI")
    time.sleep(2)
    voice_commands("Recent Applications", "Aplikasi Terbaru")
    time.sleep(2)
    voice_commands("Dismiss All", "Dismiss all")
    time.sleep(2)
    voice_commands("My Camera", "Kameraku Saya")
    time.sleep(2)
    """
    voice_commands("Take Photo", "Memotret")
    time.sleep(10)
    voice_commands("My Files", "File Saya")
    time.sleep(2)
    voice_commands("My Photos", "Foto Saya")
    time.sleep(2)
    voice_commands("Camera", "Camera")
    time.sleep(2)
    voice_commands("Select item 1", "Pilih Nomer 1")
    time.sleep(2)
    voice_commands("Navigate Back", "Satu Tingkat Kembali")
    time.sleep(2)
    voice_commands("Back one level", "Satu Tingkat Kembali")
    time.sleep(2)
    voice_commands("My files", "FILE SAYA")
    time.sleep(2)
    voice_commands("Navigate Back", "Satu Tingkat Kembali")
    time.sleep(3)
    voice_commands("Show Help", "Minta Bantuan")
    time.sleep(3)
    voice_commands("Hide Help", "Sembunyikan Bantuan")
    time.sleep(3)
    voice_commands("Flash On", "nyalakan flash")
    time.sleep(3)
    voice_commands("Flash Off", "matikan flash")
    time.sleep(3)
    voice_commands("Flash Auto", "flash otomatis")
    time.sleep(3)
    #voice_commands("Manual Focus", "fokus manual")
    #time.sleep(3)
    #voice_commands("Auto Focus", "fokus otomatis")
    #time.sleep(3)
    voice_commands("Exposure level plus 1", "tingkat pajanan plus 1")
    time.sleep(3)
    voice_commands("Exposure level plus 2", "tingkat pajanan plus 2")
    time.sleep(3)
    voice_commands("Exposure level minus 1", "tingkat pajanan minus 1")
    time.sleep(3)
    voice_commands("Exposure level minus 2", "tingkat pajanan minus 2")
    time.sleep(3)
    voice_commands("Exposure level 0", "tingkat pajanan 0")
    time.sleep(3)
    voice_commands("zoom level 2", "Tingkat Zoom 2")
    time.sleep(3)
    voice_commands("zoom level 3", "Tingkat Zoom 3")
    time.sleep(3)
    voice_commands("zoom level 4", "Tingkat Zoom 4")
    time.sleep(3)
    voice_commands("zoom level 5", "Tingkat Zoom 5")
    time.sleep(3)
    voice_commands("zoom level 1", "Tingkat Zoom 1")
    time.sleep(3)
    voice_commands("More options", "Pilihan Lain")
    time.sleep(3)
    voice_commands("Aspect ratio", "Rasio Aspek")
    time.sleep(1)
    voice_commands("4 by 3", "4 kali 3")
    time.sleep(2)
    voice_commands("Aspect ratio", "Rasio Aspek")
    time.sleep(1)
    voice_commands("16 by 9", "16 kali 9")
    time.sleep(3)
    voice_commands("Field of view", "bidang pandang")
    time.sleep(3)
    #voice_commands("Image Resolution", "resolusi gambar")
    #time.sleep(2)
    #voice_commands("Low", "rendah")
    #time.sleep(2)
    #voice_commands("Image Resolution", "resolusi gambar")
    #time.sleep(2)
    #voice_commands("High", "tinggi")
    #time.sleep(2)
    voice_commands("Video Resolution", "resolusi video")
    time.sleep(2)
    voice_commands("Low", "rendah")
    time.sleep(2)
    voice_commands("Video Resolution", "resolusi video")
    time.sleep(2)
    voice_commands("High", "tinggi")
    time.sleep(2)
    voice_commands("Frame rate", "tingkat bingkai")
    time.sleep(2)
    voice_commands("15", "15")
    time.sleep(2)
    #voice_commands("Frame rate", "tingkat bingkai")
    #time.sleep(2)
    #voice_commands("30", "30")
    #time.sleep(2)
    voice_commands("Frame rate", "tingkat bingkai")
    time.sleep(2)
    voice_commands("25", "25")
    time.sleep(4)
    voice_commands("Video Stabilization", "video stabilisasi")
    time.sleep(2)
    voice_commands("Hide Options", "Sembunyikan Pilihan")
    time.sleep(2)
    voice_commands("Record Video", "rekam video")
    time.sleep(10)
    voice_commands("Stop Recording", "hentikan rekaman")
    time.sleep(2)
    voice_commands("Preview", "PRATINJAU")
    time.sleep(2)
    voice_commands("Video Pause", "jeda video")
    time.sleep(2)
    voice_commands("Video Play", "putar video")
    time.sleep(2)
    voice_commands("Video Rewind", "putar balik video")
    time.sleep(2)
    voice_commands("Video Forward", "majukan video")
    time.sleep(2)
    voice_commands("Video Stop", "hentikan video")
    time.sleep(2)
    voice_commands("Navigate Home", "Kembali Ke Awal")
    time.sleep(2)

    workbook.close()
