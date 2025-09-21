#include <Wire.h>
#include <WiFi.h>
#include "esp_wifi.h"
#include "AD5933.h"
#include <math.h>

AD5933 ad5933;

// GPIO Pin Definitions
#define X_AXIS_ADDR_0 7
#define X_AXIS_ADDR_1 6
#define X_AXIS_ADDR_2 5
#define X_AXIS_ADDR_3 4
#define X_AXIS_ADDR_4 3
#define X_AXIS_ADDR_5 2
#define X_AXIS_ADDR_6 1

#define Y_AXIS_ADDR_0 41
#define Y_AXIS_ADDR_1 40
#define Y_AXIS_ADDR_2 39
#define Y_AXIS_ADDR_3 38
#define Y_AXIS_ADDR_4 37
#define Y_AXIS_ADDR_5 36
#define Y_AXIS_ADDR_6 35

#define ANALOG_MUX_SWITCH_0 10
#define ANALOG_MUX_SWITCH_1 11

#define DIGITAL_MUX_SWITCH_0 16
#define DIGITAL_MUX_SWITCH_1 17

#define MUX_SWITCH_ADG849 9

#define AD5933_SDA 12  // SDA Pin Number
#define AD5933_SCL 13  // SCL Pin Number

// Initial Values
int startFreq = 50000;
int frequencyUnit = 1000;
int numIncrements = 20;
int refResist = 100000;

double *gain;
double *phase;

const unsigned long ACK_TIMEOUT = 60000; // Wait for 60 seconds

void setup() {
  // Disable Wi-Fi
  WiFi.disconnect(true);
  delay(100);
  esp_wifi_stop();

  // GPIO Setup
  pinMode(X_AXIS_ADDR_0, OUTPUT); pinMode(X_AXIS_ADDR_1, OUTPUT);
  pinMode(X_AXIS_ADDR_2, OUTPUT); pinMode(X_AXIS_ADDR_3, OUTPUT);
  pinMode(X_AXIS_ADDR_4, OUTPUT); pinMode(X_AXIS_ADDR_5, OUTPUT);
  pinMode(X_AXIS_ADDR_6, OUTPUT);

  pinMode(Y_AXIS_ADDR_0, OUTPUT); pinMode(Y_AXIS_ADDR_1, OUTPUT);
  pinMode(Y_AXIS_ADDR_2, OUTPUT); pinMode(Y_AXIS_ADDR_3, OUTPUT);
  pinMode(Y_AXIS_ADDR_4, OUTPUT); pinMode(Y_AXIS_ADDR_5, OUTPUT);
  pinMode(Y_AXIS_ADDR_6, OUTPUT);

  pinMode(ANALOG_MUX_SWITCH_0, OUTPUT); pinMode(ANALOG_MUX_SWITCH_1, OUTPUT);
  pinMode(DIGITAL_MUX_SWITCH_0, OUTPUT); pinMode(DIGITAL_MUX_SWITCH_1, OUTPUT);
  pinMode(MUX_SWITCH_ADG849, OUTPUT);

  // Initialize I2C
  Wire.begin(AD5933_SDA, AD5933_SCL);

  // Initialize Serial Communication
  Serial.begin(115200);
  delay(2000); // Wait for serial communication to stabilize
  Serial.println("AD5933 Test Start");

  // Initialize AD5933
  if (!(AD5933::reset() && AD5933::setInternalClock(true) &&
        AD5933::setStartFrequency(startFreq) && AD5933::setIncrementFrequency(frequencyUnit) &&
        AD5933::setNumberIncrements(numIncrements) && AD5933::setPGAGain(PGA_GAIN_X1))) {
    Serial.println("FAILED in initialization!");
    while (true);
  }

  // Set MUX_SWITCH_ADG849 to LOW before initial calibration
  digitalWrite(MUX_SWITCH_ADG849, LOW);
  initialCalibration();
}

void loop() {
  ModeSelect();
}

void flushSerialBuffer() {
  while (Serial.available() > 0) {
    Serial.read();
  }
}

//
// Mode Selection
//
void ModeSelect() {
  int mychoice = 0;
  while (true) {
    // PROMPT: This line is detected by Python to wait for user input.
    Serial.print("Set AD5933 Mode (0: Calibration, 1: COB Impedance Measurement, 2: Rcal Impedance Measurement, 3: Diagonal Sweep, 4: COB Range Sweep, 5: Range Step Sweep): ");
    flushSerialBuffer();
    delay(10);
    while (Serial.available() == 0) { }
    mychoice = Serial.readStringUntil('\n').toInt();

    // INFO: These lines are for information only and are not detected as prompts.
    if (mychoice == 0) {
      Serial.println("Starting Calibration.");
      digitalWrite(MUX_SWITCH_ADG849, LOW);
      initialCalibration();
    } else if (mychoice == 1) {
      Serial.println("Checking impedance of COB.");
      digitalWrite(MUX_SWITCH_ADG849, HIGH);
      impedanceMeasurementCOB();
    } else if (mychoice == 2) {
      Serial.println("Checking impedance at Rcal position.");
      digitalWrite(MUX_SWITCH_ADG849, LOW);
      impedanceMeasurementRcal();
    } else if (mychoice == 3) {
      Serial.println("Starting Diagonal Sweep.");
      digitalWrite(MUX_SWITCH_ADG849, HIGH);
      impedanceMeasurementDiag();
    } else if (mychoice == 4) {
      Serial.println("Starting COB Range Sweep (7-bit input method).");
      digitalWrite(MUX_SWITCH_ADG849, HIGH);
      impedanceMeasurementCOBRange();
    } else if (mychoice == 5) {
      Serial.println("Starting COB Range Step Sweep (X/Y increment setting).");
      digitalWrite(MUX_SWITCH_ADG849, HIGH);
      impedanceMeasurementCOBRangeWithSteps();
    } else {
      Serial.println("Invalid input. Please enter 0, 1, 2, 3, 4, or 5.");
    }
  }
}

void initialCalibration() {
  showSweepMenu(); // Input sweep settings
  Serial.println("[INFO] Performing calibration."); // INFO
  gain = new double[numIncrements + 1];
  phase = new double[numIncrements + 1];

  int *real = new int[numIncrements + 1];
  int *imag = new int[numIncrements + 1];

  if (AD5933::calibrate(gain, phase, real, imag, refResist, numIncrements + 1)) {
    Serial.println("[INFO] Calibration complete!"); // INFO
    Serial.println("=================================================================================================================================");
    for (int i = 0; i <= numIncrements; i++) {
      double magnitude = sqrt(pow(real[i], 2) + pow(imag[i], 2));
      Serial.print("Cal Point ");
      Serial.print(i);
      Serial.print(": R=");
      Serial.print(real[i]);
      Serial.print(" / I=");
      Serial.print(imag[i]);
      Serial.print("\t |Z|=");
      Serial.print(magnitude);
      Serial.print("\t System Phase=");
      Serial.print(phase[i]);
      Serial.println(" degrees");
    }
    Serial.println("=================================================================================================================================");
  } else {
    Serial.println("[ERROR] Calibration failed..."); // INFO
  }

  delete[] real;
  delete[] imag;
}

// 1. COB Single Position Sweep
void impedanceMeasurementCOB() {
  setMuxGroup();
  getAddressInput();
  frequencySweepRaw(startFreq, frequencyUnit, numIncrements);
}

// 2. Rcal Sweep
void impedanceMeasurementRcal() {
  frequencySweepRaw(startFreq, frequencyUnit, numIncrements);
}

// 3. Diagonal Sweep (plotting function currently removed)
void impedanceMeasurementDiag() {
  setMuxGroup();
  diagonalSweepPattern();
}

// 4. COB Range Sweep
void impedanceMeasurementCOBRange() {
  setMuxGroup();
  sweepCOBRange();
}

// 5. COB Range Step Sweep
void impedanceMeasurementCOBRangeWithSteps() {
  setMuxGroup();
  sweepCOBRangeWithSteps();
}

//
// User Input and Sweep Settings
//
void showSweepMenu() {
  // PROMPT: Input Start Frequency
  while (true) {
    Serial.print("Enter the start frequency (1~100 kHz): ");
    flushSerialBuffer();
    delay(10);
    while (Serial.available() == 0) { }
    startFreq = Serial.readStringUntil('\n').toInt();
    if (startFreq >= 1 && startFreq <= 100) {
      startFreq *= 1000;
      Serial.print("[INFO] Set start frequency: "); // INFO
      Serial.print(startFreq);
      Serial.println(" Hz");
      break;
    } else {
      Serial.println("[ERROR] Invalid input. Please enter a value between 1 and 100 kHz.");
    }
  }
  // PROMPT: Input Frequency Increment
  while (true) {
    Serial.print("Enter the frequency increment (1~10000 Hz): ");
    flushSerialBuffer();
    delay(10);
    while (Serial.available() == 0) { }
    frequencyUnit = Serial.readStringUntil('\n').toInt();
    if (frequencyUnit >= 1 && frequencyUnit <= 10000) {
      Serial.print("[INFO] Set frequency increment: "); // INFO
      Serial.print(frequencyUnit);
      Serial.println(" Hz");
      break;
    } else {
      Serial.println("[ERROR] Invalid input. Please enter a value between 1 and 10000 Hz.");
    }
  }
  // PROMPT: Input Number of Increments
  while (true) {
    Serial.print("Enter the number of measurements (1~100): ");
    flushSerialBuffer();
    delay(10);
    while (Serial.available() == 0) { }
    numIncrements = Serial.readStringUntil('\n').toInt();
    if (numIncrements >= 1 && numIncrements <= 100) {
      Serial.print("[INFO] Set number of measurements: "); // INFO
      Serial.print(numIncrements);
      Serial.println(" times");
      break;
    } else {
      Serial.println("[ERROR] Invalid input. Please enter a value between 1 and 100.");
    }
  }
  // PROMPT: Input Settling Time Cycles
  while (true) {
    Serial.print("Enter Settling Time Cycles (0~511): ");
    flushSerialBuffer();
    delay(10);
    while (Serial.available() == 0) { }
    int settlingTimeCycles = Serial.readStringUntil('\n').toInt();
    if (settlingTimeCycles >= 0 && settlingTimeCycles <= 511) {
      ad5933.setSettlingCycles(settlingTimeCycles);
      Serial.print("[INFO] Set Settling Time Cycles: "); // INFO
      Serial.println(settlingTimeCycles);
      break;
    } else {
      Serial.println("[ERROR] Invalid input. Please enter a value between 0 and 511.");
    }
  }
  // PROMPT: Set Output Excitation Range
  while (true) {
    Serial.print("Select Output Excitation Range (1: 2 Vpp, 2: 1 Vpp, 3: 0.4 Vpp, 4: 0.2 Vpp): ");
    flushSerialBuffer();
    delay(10);
    while (Serial.available() == 0) { }
    int excitationRange = Serial.readStringUntil('\n').toInt();
    switch (excitationRange) {
      case 1: ad5933.setRange(CTRL_OUTPUT_RANGE_1); Serial.println("[INFO] Set to 2 Vpp (Range 1)."); break;
      case 2: ad5933.setRange(CTRL_OUTPUT_RANGE_2); Serial.println("[INFO] Set to 1 Vpp (Range 2)."); break;
      case 3: ad5933.setRange(CTRL_OUTPUT_RANGE_3); Serial.println("[INFO] Set to 0.4 Vpp (Range 3)."); break;
      case 4: ad5933.setRange(CTRL_OUTPUT_RANGE_4); Serial.println("[INFO] Set to 0.2 Vpp (Range 4)."); break;
      default: Serial.println("[ERROR] Invalid input. Please select one from 1-4."); continue;
    }
    break;
  }
  // PROMPT: Set PGA Control
  while (true) {
    Serial.print("Select PGA Gain (1 or 5): ");
    flushSerialBuffer();
    delay(10);
    while (Serial.available() == 0) { }
    int pgaControl = Serial.readStringUntil('\n').toInt();
    if (pgaControl == 1 || pgaControl == 5) {
      AD5933::setPGAGain(pgaControl == 1 ? PGA_GAIN_X1 : PGA_GAIN_X5);
      Serial.print("[INFO] PGA Gain set to: x"); // INFO
      Serial.println(pgaControl);
      break;
    } else {
      Serial.println("[ERROR] Invalid input. Please enter 1 or 5.");
    }
  }
  // PROMPT: Input Calibration Impedance
  while (true) {
    Serial.print("Enter Calibration Impedance (in Ohms, positive integer): ");
    flushSerialBuffer();
    delay(10);
    while (Serial.available() == 0) { }
    refResist = Serial.readStringUntil('\n').toInt();
    if (refResist > 0) {
      Serial.print("[INFO] Set Calibration Impedance: "); // INFO
      Serial.print(refResist);
      Serial.println(" ohm");
      break;
    } else {
      Serial.println("[ERROR] Invalid input. Please enter a valid Calibration Impedance value.");
    }
  }
  // Set AD5933 Frequency Sweep
  if (!(AD5933::setStartFrequency(startFreq) &&
        AD5933::setIncrementFrequency(frequencyUnit) &&
        AD5933::setNumberIncrements(numIncrements))) {
    Serial.println("[ERROR] Failed to set frequency sweep.");
  } else {
    Serial.println("[INFO] Frequency sweep settings complete.");
  }
}

// Input for a single coordinate
void getAddressInput() {
  int xAddress[7];
  int yAddress[7];
  const int X_AXIS_PINS[7] = {X_AXIS_ADDR_0, X_AXIS_ADDR_1, X_AXIS_ADDR_2, X_AXIS_ADDR_3, X_AXIS_ADDR_4, X_AXIS_ADDR_5, X_AXIS_ADDR_6};
  const int Y_AXIS_PINS[7] = {Y_AXIS_ADDR_0, Y_AXIS_ADDR_1, Y_AXIS_ADDR_2, Y_AXIS_ADDR_3, Y_AXIS_ADDR_4, Y_AXIS_ADDR_5, Y_AXIS_ADDR_6};

  // INFO: This is an instruction line, not a prompt for python.
  Serial.println("Instructions: Enter X-axis Address (7 digits, each bit as 0 or 1):");
  for (int i = 0; i < 7; i++) {
    while (true) {
      // PROMPT: This line starts with "X Axis Address" and is detected by Python.
      Serial.print("X Axis Address ");
      Serial.print(i);
      Serial.print(" (0 or 1): ");
      while (Serial.available() == 0) { }
      char c = Serial.read();
      flushSerialBuffer();
      if (c == '0' || c == '1') {
        int bitValue = c - '0';
        xAddress[i] = bitValue;
        digitalWrite(X_AXIS_PINS[i], bitValue == 1 ? HIGH : LOW);
        break;
      } else {
        Serial.println("[ERROR] Invalid input. Enter 0 or 1.");
        flushSerialBuffer();
      }
    }
  }
  String xAddrStr = "";
  for (int i = 0; i < 7; i++) {
    xAddrStr += String(xAddress[i]);
  }
  
  // INFO: This is an instruction line, not a prompt for python.
  Serial.println("Instructions: Enter Y-axis Address (7 digits, each bit as 0 or 1):");
  for (int i = 0; i < 7; i++) {
    while (true) {
      // PROMPT: This line starts with "Y Axis Address" and is detected by Python.
      Serial.print("Y Axis Address ");
      Serial.print(i);
      Serial.print(" (0 or 1): ");
      while (Serial.available() == 0) { }
      char c = Serial.read();
      flushSerialBuffer();
      if (c == '0' || c == '1') {
        int bitValue = c - '0';
        yAddress[i] = bitValue;
        digitalWrite(Y_AXIS_PINS[i], bitValue == 1 ? HIGH : LOW);
        break;
      } else {
        Serial.println("[ERROR] Invalid input. Enter 0 or 1.");
        flushSerialBuffer();
      }
    }
  }
  String yAddrStr = "";
  for (int i = 0; i < 7; i++) {
    yAddrStr += String(yAddress[i]);
  }
  Serial.print("[INFO] Set X-axis Address : "); // INFO
  Serial.print(xAddrStr);
  Serial.print(", Y-axis Address : ");
  Serial.println(yAddrStr);
}

// Diagonal Sweep
void diagonalSweepPattern() {
  const int X_AXIS_PINS[7] = {X_AXIS_ADDR_0, X_AXIS_ADDR_1, X_AXIS_ADDR_2, X_AXIS_ADDR_3, X_AXIS_ADDR_4, X_AXIS_ADDR_5, X_AXIS_ADDR_6};
  const int Y_AXIS_PINS[7] = {Y_AXIS_ADDR_0, Y_AXIS_ADDR_1, Y_AXIS_ADDR_2, Y_AXIS_ADDR_3, Y_AXIS_ADDR_4, Y_AXIS_ADDR_5, Y_AXIS_ADDR_6};

  Serial.println("[INFO] Starting forward diagonal sweep...");
  for (int i = 0; i < 7; i++) {
    for (int j = 0; j < 7; j++) {
      digitalWrite(X_AXIS_PINS[j], i == j ? HIGH : LOW);
      digitalWrite(Y_AXIS_PINS[j], i == j ? HIGH : LOW);
    }
    delay(100);
    Serial.print("[INFO] Current X address: ");
    for (int k = 0; k < 7; k++) Serial.print(digitalRead(X_AXIS_PINS[k]));
    Serial.print(" | Y address: ");
    for (int k = 0; k < 7; k++) Serial.print(digitalRead(Y_AXIS_PINS[k]));
    Serial.println();
    frequencySweepRaw(startFreq, frequencyUnit, numIncrements);
  }
  Serial.println("[INFO] Starting reverse diagonal sweep...");
  for (int i = 6; i >= 0; i--) {
    for (int j = 0; j < 7; j++) {
      digitalWrite(X_AXIS_PINS[j], i == j ? HIGH : LOW);
      digitalWrite(Y_AXIS_PINS[j], i == j ? HIGH : LOW);
    }
    delay(100);
    Serial.print("[INFO] Current X address: ");
    for (int k = 0; k < 7; k++) Serial.print(digitalRead(X_AXIS_PINS[k]));
    Serial.print(" | Y address: ");
    for (int k = 0; k < 7; k++) Serial.print(digitalRead(Y_AXIS_PINS[k]));
    Serial.println();
    frequencySweepRaw(startFreq, frequencyUnit, numIncrements);
  }
  Serial.println("[INFO] Diagonal sweep complete.");
}

//
// Range Sweep Related Functions
//
void sweepCOBRange() {
  Serial.println("[INFO] Starting COB range sweep (fixed step size of 1)...");
  int xStart, xEnd, yStart, yEnd;

  while (true) {
    xStart = readSingleAddress("Instructions: Enter X-axis start address (7-bit binary):");
    xEnd   = readSingleAddress("Instructions: Enter X-axis end address (7-bit binary):");
    if (xStart > xEnd) {
      Serial.println("[ERROR] X-axis start address is greater than end address. Please re-enter.");
      continue;
    }
    Serial.print("[INFO] Entered X-axis range: Start = ");
    Serial.print(intToBinaryString(xStart));
    Serial.print(", End = ");
    Serial.println(intToBinaryString(xEnd));
    // PROMPT
    Serial.print("Is this range correct? (Y/N): ");
    flushSerialBuffer();
    while (Serial.available() == 0) { }
    char confirm = Serial.read();
    Serial.println(confirm);
    flushSerialBuffer();
    if (confirm == 'Y' || confirm == 'y') break;
    else Serial.println("[INFO] Re-entering X-axis range.");
  }

  while (true) {
    yStart = readSingleAddress("Instructions: Enter Y-axis start address (7-bit binary):");
    yEnd   = readSingleAddress("Instructions: Enter Y-axis end address (7-bit binary):");
    if (yStart > yEnd) {
      Serial.println("[ERROR] Y-axis start address is greater than end address. Please re-enter.");
      continue;
    }
    Serial.print("[INFO] Entered Y-axis range: Start = ");
    Serial.print(intToBinaryString(yStart));
    Serial.print(", End = ");
    Serial.println(intToBinaryString(yEnd));
    // PROMPT
    Serial.print("Is this range correct? (Y/N): ");
    flushSerialBuffer();
    while (Serial.available() == 0) { }
    char confirm = Serial.read();
    Serial.println(confirm);
    flushSerialBuffer();
    if (confirm == 'Y' || confirm == 'y') break;
    else Serial.println("[INFO] Re-entering Y-axis range.");
  }

  const int X_AXIS_PINS[7] = {X_AXIS_ADDR_0, X_AXIS_ADDR_1, X_AXIS_ADDR_2, X_AXIS_ADDR_3, X_AXIS_ADDR_4, X_AXIS_ADDR_5, X_AXIS_ADDR_6};
  const int Y_AXIS_PINS[7] = {Y_AXIS_ADDR_0, Y_AXIS_ADDR_1, Y_AXIS_ADDR_2, Y_AXIS_ADDR_3, Y_AXIS_ADDR_4, Y_AXIS_ADDR_5, Y_AXIS_ADDR_6};

  for (int x = xStart; x <= xEnd; x++) {
    for (int y = yStart; y <= yEnd; y++) {
      for (int i = 0; i < 7; i++) { int bitVal = (x >> (6 - i)) & 1; digitalWrite(X_AXIS_PINS[i], bitVal ? HIGH : LOW); }
      for (int i = 0; i < 7; i++) { int bitVal = (y >> (6 - i)) & 1; digitalWrite(Y_AXIS_PINS[i], bitVal ? HIGH : LOW); }
      
      Serial.print("Current_Coord->X=");
      Serial.print(intToBinaryString(x));
      Serial.print(",Y=");
      Serial.println(intToBinaryString(y));

      Serial.println("SWEEP_START");
      frequencySweepRaw(startFreq, frequencyUnit, numIncrements);
      Serial.println("SWEEP_DONE");

      unsigned long startTime = millis();
      bool storeOKReceived = false;
      while (millis() - startTime < ACK_TIMEOUT) {
        if (Serial.available() > 0) {
          String ackLine = Serial.readStringUntil('\n');
          if (ackLine.indexOf("STORE_OK") != -1) {
            storeOKReceived = true;
            break;
          }
        }
      }
      if (!storeOKReceived) {
        Serial.println("[ERROR] Data save failed. Retrying measurement.");
        frequencySweepRaw(startFreq, frequencyUnit, numIncrements);
        Serial.println("SWEEP_DONE");
        unsigned long retryStart = millis();
        bool storeOKReceivedRetry = false;
        while (millis() - retryStart < ACK_TIMEOUT) {
          if (Serial.available() > 0) {
            String ackLine2 = Serial.readStringUntil('\n');
            if (ackLine2.indexOf("STORE_OK") != -1) {
              storeOKReceivedRetry = true;
              break;
            }
          }
        }
        if(!storeOKReceivedRetry) {
          Serial.println("[ERROR] Retried data save failed. Moving to the next coordinate.");
        }
      }
      delay(200);
    }
  }
  Serial.println("[INFO] COB range sweep complete.");
}

//
// Range + Step Sweep Function (Mode 5)
//
void sweepCOBRangeWithSteps() {
  Serial.println("[INFO] Starting COB range sweep (step increment mode, boundaries included)...");
  int xStart, xEnd, yStart, yEnd, xStep, yStep;

  while (true) {
    xStart = readSingleAddress("Instructions: Enter X-axis start address (7-bit binary):");
    xEnd   = readSingleAddress("Instructions: Enter X-axis end address (7-bit binary):");
    if (xStart > xEnd) {
      Serial.println("[ERROR] X-axis start address is greater than end address. Please re-enter.");
      continue;
    }
    Serial.print("[INFO] Entered X-axis range: Start = ");
    Serial.print(intToBinaryString(xStart));
    Serial.print(", End = ");
    Serial.println(intToBinaryString(xEnd));

    while (true) {
      Serial.print("Enter X-axis increment unit (1~127): "); // PROMPT
      flushSerialBuffer();
      while (Serial.available() == 0) { }
      xStep = Serial.readStringUntil('\n').toInt();
      if (xStep < 1 || xStep > 127) { Serial.println("[ERROR] Invalid input. Please enter a value between 1 and 127."); } 
      else { break; }
    }

    Serial.println();
    yStart = readSingleAddress("Instructions: Enter Y-axis start address (7-bit binary):");
    yEnd   = readSingleAddress("Instructions: Enter Y-axis end address (7-bit binary):");
    if (yStart > yEnd) {
      Serial.println("[ERROR] Y-axis start address is greater than end address. Please re-enter.");
      continue;
    }
    Serial.print("[INFO] Entered Y-axis range: Start = ");
    Serial.print(intToBinaryString(yStart));
    Serial.print(", End = ");
    Serial.println(intToBinaryString(yEnd));

    while (true) {
      Serial.print("Enter Y-axis increment unit (1~127): "); // PROMPT
      flushSerialBuffer();
      while (Serial.available() == 0) { }
      yStep = Serial.readStringUntil('\n').toInt();
      if (yStep < 1 || yStep > 127) { Serial.println("[ERROR] Invalid input. Please enter a value between 1 and 127."); } 
      else { break; }
    }

    Serial.println("=====================================================================");
    Serial.print("X-axis range: Start="); Serial.print(intToBinaryString(xStart)); Serial.print(", End="); Serial.print(intToBinaryString(xEnd)); Serial.print(", Increment="); Serial.println(xStep);
    Serial.print("Y-axis range: Start="); Serial.print(intToBinaryString(yStart)); Serial.print(", End="); Serial.print(intToBinaryString(yEnd)); Serial.print(", Increment="); Serial.println(yStep);
    Serial.println("=====================================================================");
    Serial.print("Is this range correct? (Y/N): "); // PROMPT
    flushSerialBuffer();
    while (Serial.available() == 0) { }
    char confirm = Serial.read();
    Serial.println(confirm);
    flushSerialBuffer();

    if (confirm == 'Y' || confirm == 'y') { break; } 
    else { Serial.println("[INFO] Re-entering range."); }
  }

  const int X_AXIS_PINS[7] = {X_AXIS_ADDR_0, X_AXIS_ADDR_1, X_AXIS_ADDR_2, X_AXIS_ADDR_3, X_AXIS_ADDR_4, X_AXIS_ADDR_5, X_AXIS_ADDR_6};
  const int Y_AXIS_PINS[7] = {Y_AXIS_ADDR_0, Y_AXIS_ADDR_1, Y_AXIS_ADDR_2, Y_AXIS_ADDR_3, Y_AXIS_ADDR_4, Y_AXIS_ADDR_5, Y_AXIS_ADDR_6};

  int xVal = xStart;
  while (true) {
    if (xVal > xEnd) xVal = xEnd;
    int yVal = yStart;
    while (true) {
      if (yVal > yEnd) yVal = yEnd;

      for (int i = 0; i < 7; i++) { int bitVal = (xVal >> (6 - i)) & 1; digitalWrite(X_AXIS_PINS[i], bitVal ? HIGH : LOW); }
      for (int i = 0; i < 7; i++) { int bitVal = (yVal >> (6 - i)) & 1; digitalWrite(Y_AXIS_PINS[i], bitVal ? HIGH : LOW); }

      Serial.print("Current_Coord->X="); Serial.print(intToBinaryString(xVal)); Serial.print(",Y="); Serial.println(intToBinaryString(yVal));
      Serial.println("SWEEP_START");
      frequencySweepRaw(startFreq, frequencyUnit, numIncrements);
      Serial.println("SWEEP_DONE");
      
      unsigned long startTime = millis();
      bool storeOKReceived = false;
      while (millis() - startTime < ACK_TIMEOUT) {
        if (Serial.available() > 0) {
          String ackLine = Serial.readStringUntil('\n');
          if (ackLine.indexOf("STORE_OK") != -1) { storeOKReceived = true; break; }
        }
      }
      if (!storeOKReceived) {
        Serial.println("[ERROR] Data save failed. Retrying measurement.");
        frequencySweepRaw(startFreq, frequencyUnit, numIncrements);
        Serial.println("SWEEP_DONE");
        unsigned long retryStart = millis();
        bool storeOKReceivedRetry = false;
        while (millis() - retryStart < ACK_TIMEOUT) {
          if (Serial.available() > 0) {
            String ackLine2 = Serial.readStringUntil('\n');
            if (ackLine2.indexOf("STORE_OK") != -1) { storeOKReceivedRetry = true; break; }
          }
        }
        if(!storeOKReceivedRetry) { Serial.println("[ERROR] Retried data save failed. Moving to next coordinate."); }
      }
      delay(200);

      if (yVal == yEnd) break;
      yVal += yStep;
    }
    if (xVal == xEnd) break;
    xVal += xStep;
  }

  Serial.println("[INFO] COB range step sweep complete.");
}

//
// readSingleAddress(): Function to read a 7-bit address and convert it to an integer
//
int readSingleAddress(String prompt) {
  String binStr = "";
  int address = 0;
  // INFO: Displays instructions passed from the calling function.
  Serial.println(prompt);
  for (int i = 0; i < 7; i++) {
    while (true) {
      // PROMPT: This line starts with "  Bit" and is detected by Python.
      Serial.print("  Bit ");
      Serial.print(i);
      Serial.print(" (0 or 1): ");
      flushSerialBuffer();
      while (Serial.available() == 0) { }
      char c = Serial.read();
      flushSerialBuffer();
      if (c == '0' || c == '1') {
        address = (address << 1) | (c - '0');
        binStr += c;
        break;
      } else {
        Serial.println("  [ERROR] Invalid input. Please enter 0 or 1.");
      }
    }
  }
  Serial.print("[INFO] Address entered: "); // INFO
  Serial.println(binStr);
  return address;
}

//
// intToBinaryString(): Function to return an integer as a 7-bit binary string
//
String intToBinaryString(int number) {
  String result = "";
  for (int i = 6; i >= 0; i--) {
    result += String((number >> i) & 1);
  }
  return result;
}

//
// Set MUX Group
//
void setMuxGroup() {
  int group = 0;
  while (1) {
    // PROMPT
    Serial.print("Select MUX group (1, 2, 3, 4): ");
    flushSerialBuffer();
    delay(10);
    while (Serial.available() == 0) { }
    String inputStr = Serial.readStringUntil('\n');
    group = inputStr.toInt();
    if (group >= 1 && group <= 4) {
      Serial.print("[INFO] Group "); // INFO
      Serial.print(group);
      Serial.println(" selected");
      break;
    } else {
      Serial.println("[ERROR] Invalid input. Please enter a value between 1 and 4.");
    }
  }
  switch (group) {
    case 1: /* 00 */ digitalWrite(ANALOG_MUX_SWITCH_0, LOW); digitalWrite(ANALOG_MUX_SWITCH_1, LOW); digitalWrite(DIGITAL_MUX_SWITCH_0, LOW); digitalWrite(DIGITAL_MUX_SWITCH_1, LOW); break;
    case 2: /* 01 */ digitalWrite(ANALOG_MUX_SWITCH_0, HIGH); digitalWrite(ANALOG_MUX_SWITCH_1, LOW); digitalWrite(DIGITAL_MUX_SWITCH_0, HIGH); digitalWrite(DIGITAL_MUX_SWITCH_1, LOW); break;
    case 3: /* 10 */ digitalWrite(ANALOG_MUX_SWITCH_0, LOW); digitalWrite(ANALOG_MUX_SWITCH_1, HIGH); digitalWrite(DIGITAL_MUX_SWITCH_0, LOW); digitalWrite(DIGITAL_MUX_SWITCH_1, HIGH); break;
    case 4: /* 11 */ digitalWrite(ANALOG_MUX_SWITCH_0, HIGH); digitalWrite(ANALOG_MUX_SWITCH_1, HIGH); digitalWrite(DIGITAL_MUX_SWITCH_0, HIGH); digitalWrite(DIGITAL_MUX_SWITCH_1, HIGH); break;
  }
  Serial.println("[INFO] MUX switches have been set."); // INFO
}

//
// Frequency Sweep (Outputs measurement results)
//
void frequencySweepRaw(int startFreq, int frequencyUnit, int numIncrements) {
  int real, imag, i = 0;
  double cfreq = startFreq / 1000.0;

  if (!(AD5933::setPowerMode(POWER_STANDBY) &&
        AD5933::setControlMode(CTRL_INIT_START_FREQ) &&
        AD5933::setControlMode(CTRL_START_FREQ_SWEEP))) {
    Serial.println("[ERROR] Could not initialize frequency sweep...");
    return;
  }

  Serial.println("=================================================================================================================================");

  while ((AD5933::readStatusRegister() & STATUS_SWEEP_DONE) != STATUS_SWEEP_DONE) {
    if (!AD5933::getComplexData(&real, &imag)) {
      Serial.println("[ERROR] Could not get raw frequency data...");
      real = 0;
      imag = 0;
    }

    double magnitude = sqrt(pow(real, 2) + pow(imag, 2));
    double impedance = 1 / (magnitude * gain[i]); // Calculate calibrated impedance
    double rawPhase = atan2(imag, real) * (180.0 / M_PI);

    if (rawPhase < 0) { rawPhase += 360.0; }
    double correctedPhase = rawPhase - phase[i];

    if (correctedPhase < -180.0) { correctedPhase += 360.0; } 
    else if (correctedPhase >= 180.0) { correctedPhase -= 360.0; }

    double Rreal = impedance * cos(correctedPhase * M_PI / 180.0);
    double Ximaginary = impedance * sin(correctedPhase * M_PI / 180.0);

    Serial.print(cfreq);
    Serial.print("kHz: R=");
    Serial.print(real);
    Serial.print("/I=");
    Serial.print(imag);
    Serial.print("\t  |Z|=");
    Serial.print(impedance);
    Serial.print("\t  Phase=");
    Serial.print(correctedPhase);
    Serial.print(" degrees\t Resistance=");
    Serial.print(Rreal);
    Serial.print("\t Reactance=");
    Serial.println(Ximaginary);

    i++;
    cfreq += frequencyUnit / 1000.0;
    AD5933::setControlMode(CTRL_INCREMENT_FREQ);
  }

  Serial.println("Frequency sweep complete!");
  Serial.println("=================================================================================================================================");

  if (!AD5933::setPowerMode(POWER_STANDBY)) {
    Serial.println("[ERROR] Could not set to standby...");
  }
}