#include "HX711.h"
#include <Wire.h>
#include <LiquidCrystal_I2C.h>

// Definición de los pines del sensor HX711
const int pinOUT = A0;  // Pin de datos
const int pinSCK = A1;  // Pin de reloj

// Creación del objeto sensor
HX711 sensor;

// Configuración del LCD
LiquidCrystal_I2C lcd(0x27, 16, 4);  // Ajusta la dirección 0x27 según tu módulo

unsigned long tiempoInicio = 0;
bool medicionIniciada = false;
bool buscandoDiastolica = false;
unsigned long tiempoSistolica = 0;

float presionSistolica = 0.0;
float presionDiastolica = 0.0; 

void setup() {
  Serial.begin(9600);       // Comunicación serial a través de USB
  Serial1.begin(9600);      // Comunicación serial a través de Bluetooth (HC-05)

  sensor.begin(pinOUT, pinSCK);
  sensor.set_scale();
  sensor.tare();

  // Inicializar el LCD
  lcd.init();
  lcd.backlight();
  lcd.clear();

  // Mostrar mensaje de bienvenida
    lcd.setCursor(0, 0);
  mostrarEnLCD("Bienvenido");
}

void loop() {
  if (Serial.available() > 0) {
    String mensajeCompleto = Serial.readStringUntil('\n');
    mostrarEnLCD(mensajeCompleto);
  }

  if (sensor.is_ready()) {
    long lecturaBruta = sensor.read();
    float factorCorreccion = 0.766 * 1.21 * 0.88; // Factor de corrección
    float lecturaMmHg = ((lecturaBruta / 16777216.0) * 300.0) * factorCorreccion;

    // Enviar datos a través de ambos seriales
    Serial.println(lecturaMmHg);
    Serial1.println(lecturaMmHg);

    if (!medicionIniciada && lecturaMmHg > 20) {
      medicionIniciada = true;
      tiempoInicio = millis();
    }

    if (medicionIniciada) {
      if (millis() - tiempoInicio <= 30000) {
        if (lecturaMmHg > presionSistolica) {
          presionSistolica = lecturaMmHg;
          tiempoSistolica = millis();
        }
      } else if (!buscandoDiastolica && millis() - tiempoSistolica >= 1000) {
        buscandoDiastolica = true;
      }

      if (buscandoDiastolica) {
        if (lecturaMmHg < presionDiastolica || presionDiastolica == 0.0) {
          presionDiastolica = lecturaMmHg;
        }
      }

      if (lecturaMmHg >= 122.35) {
        Serial.println("MAX_VALUE_REACHED");
        Serial1.println("MAX_VALUE_REACHED");
        delay(7000);

        unsigned long start = millis();
        while (millis() - start < 10000) {
          lecturaBruta = sensor.read();
          lecturaMmHg = ((lecturaBruta / 16777216.0) * 300.0) * factorCorreccion;
          Serial.println(lecturaMmHg);
          Serial1.println(lecturaMmHg);
          delay(100);
        }

        medicionIniciada = false;
        buscandoDiastolica = false;
        tiempoInicio = 0;
        presionSistolica = 0.0;
        presionDiastolica = 0.0;
      }
    }
  } else {
    Serial.println("Sensor no listo");
    Serial1.println("Sensor no listo");
  }
  delay(100);
}

void mostrarEnLCD(String mensaje) {
  lcd.clear();
  lcd.setCursor(0, 0);
  delay(5); // Da tiempo para limpiar el display

  // Divide el mensaje en partes separadas por '|'
  int primeraBarra = mensaje.indexOf('|');
  int segundaBarra = mensaje.indexOf('|', primeraBarra + 1);

  if (primeraBarra != -1 && segundaBarra != -1) {
    // Muestra la presión sistólica
    String sistolica = mensaje.substring(0, primeraBarra);
    lcd.print("Sist: " + sistolica + " mmHg");

    // Muestra la presión diastólica
    String diastolica = mensaje.substring(primeraBarra + 1, segundaBarra);
    lcd.setCursor(0, 1);
    lcd.print("Diast: " + diastolica + " mmHg");

    // Muestra el diagnóstico
    String diagnostico = mensaje.substring(segundaBarra + 1);
    lcd.setCursor(0, 2);
    lcd.print(diagnostico);
  } else {
    lcd.print(mensaje);
  }
}
