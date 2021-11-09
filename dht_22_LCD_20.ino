#include "DHT.h"
#include <Wire.h> 
#include <LiquidCrystal_I2C.h>

#define DHTPIN 7 // pino que estamos conectado
#define DHTTYPE DHT22 // DHT 22
 

DHT dht(DHTPIN, DHTTYPE);
LiquidCrystal_I2C lcd(0x27,20,4);
 
void setup() 
{
  lcd.init();                      // initialize the lcd 
  lcd.backlight();
  Serial.begin(9600);
  dht.begin();
  delay(2000);
}
 
void loop() {
  lcd.clear();
  float h = dht.readHumidity();
  float t = dht.readTemperature();
  Serial.print("Humi:");
  Serial.print(h);
  Serial.print(" ");
  Serial.print("Temp:");
  Serial.println(t);
  lcd.setCursor(4,0);
  lcd.print("Temperature:");
  lcd.setCursor(7,1);
  lcd.print(t);
  lcd.setCursor(5,2);
  lcd.print("Humidity:");
  lcd.setCursor(7,3);
  lcd.print(h);
  delay(3000);}
