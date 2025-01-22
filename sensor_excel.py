import serial
import openpyxl
from datetime import datetime

ser = serial.Serial("COM10", baudrate=9600, timeout=1) #kart bağlantı
dosya = openpyxl.load_workbook("./NEM_SENSOR.xlsx") #excel dosya açma
sayfa = dosya.active

current_time = datetime.now().strftime("%Y-%m-%d %H-%M-%S") #excel sayfa adı yıl ay gün saat
sayfa.title = current_time 

#dosya içine satır excelde en sonki satır sayısı yazılıyor.
with open("./dosya.txt","r",encoding="utf-8") as file:
    satir_sayisi=file.read()
i=int(satir_sayisi)

try:
    while True:
        data = ser.readline().decode("ascii").strip()
        if data:
            try:
                adc_value, voltage_value = data.split() 
                
                print("ADC value: ", adc_value)
                print("Voltage value: ", voltage_value)
                
                sayfa["A1"].value="ADC Value"
                sayfa["B1"].value="Voltage Value"
                sayfa[f"A{i}"].value = adc_value
                sayfa[f"B{i}"].value = voltage_value
                
                i+= 1  
                
                dosya.save("./NEM_SENSOR.xlsx")

            except ValueError:
                print("Unexpected value: ", data)

except KeyboardInterrupt:
    print("program kapatıldı")
    with open("dosya.txt","w",encoding="utf-8") as dosya:
        dosya.write(str(i))
        