#include <Wire.h>
#include <WiFi.h>
#include "esp_wifi.h"
#include "AD5933.h"

AD5933 ad5933;

// GPIO 핀 설정
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

#define AD5933_SDA 12  // SDA 핀 번호
#define AD5933_SCL 13  // SCL 핀 번호

// 초기값 설정
int startFreq = 50000;
int frequencyUnit = 1000;
int numIncrements = 20;
int refResist = 100000;

double *gain;
double *phase;

void setup() {
  // Wi-Fi 비활성화
  WiFi.disconnect(true);
  delay(100);
  esp_wifi_stop();

  // GPIO 설정
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

  // I2C 초기화
  Wire.begin(AD5933_SDA, AD5933_SCL);

  // 시리얼 통신 초기화
  Serial.begin(115200);
  delay(2000); // 시리얼 통신 안정화
  // while (!Serial);  

  Serial.println("AD5933 테스트 시작");

  // AD5933 초기화
  if (!(AD5933::reset() && AD5933::setInternalClock(true) &&
        AD5933::setStartFrequency(startFreq) && AD5933::setIncrementFrequency(frequencyUnit) &&
        AD5933::setNumberIncrements(numIncrements) && AD5933::setPGAGain(PGA_GAIN_X1))) {
    Serial.println("FAILED in initialization!");
    while (true);
  }

  // MUX_SWITCH_ADG849를 LOW로 설정하고 초기 캘리브레이션 수행
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

void ModeSelect() {
  int mychoice = 0;
  while (true) {
    Serial.print("AD5933 모드 설정 (0: 캘리브레이션 모드, 1: COB 임피던스 측정, 2: Rcal 임피던스 측정): ");
    flushSerialBuffer(); // 버퍼 비우기 추가
    delay(10);  // 시리얼 통신 안정화를 위해 대기시간 추가
    while (Serial.available() == 0) {
    }

    mychoice = Serial.readStringUntil('\n').toInt();

    if (mychoice == 0) {
      Serial.println("캘리브레이션을 시작합니다.");
      digitalWrite(MUX_SWITCH_ADG849, LOW);
      initialCalibration();
    } else if (mychoice == 1) {
      Serial.println("COB의 임피던스를 체크합니다.");
      digitalWrite(MUX_SWITCH_ADG849, HIGH); // COB 임피던스
      impedanceMeasurementCOB();
    } else if (mychoice == 2) {
      Serial.println("Rcal 위치의 임피던스를 체크합니다.");
      digitalWrite(MUX_SWITCH_ADG849, LOW); // Rcal 위치로 설정
      impedanceMeasurementRcal();
    } else {
      Serial.println("잘못된 입력. 0, 1 또는 2 입력.");
    }
  }
}

void initialCalibration() {
  showSweepMenu(); // 스윕 세팅 값을 입력받음

  Serial.println("캘리브레이션을 진행합니다.");
  gain = new double[numIncrements + 1];
  phase = new double[numIncrements + 1];

  int *real = new int[numIncrements + 1];
  int *imag = new int[numIncrements + 1];

  // 캘리브레이션 수행 및 각 포인트에 대한 데이터 수집
  if (AD5933::calibrate(gain, phase, real, imag, refResist, numIncrements + 1)) {
    Serial.println("캘리브레이션 완료!");

    Serial.println("=================================================================================================================================");

    // 각 포인트에 대해 캘리브레이션 데이터 출력
    for (int i = 0; i <= numIncrements; i++) {
      double magnitude = sqrt(pow(real[i], 2) + pow(imag[i], 2));
      Serial.print("Cal Point ");
      Serial.print(i);
      Serial.print(": R=");
      Serial.print(real[i]);
      Serial.print(" / I=");
      Serial.print(imag[i]);
      Serial.print("\t |Z|=");
      Serial.print(magnitude);  // 캘리브레이션된 임피던스 계산
      Serial.print("\t System Phase=");
      Serial.print(phase[i]);  // 시스템 위상
      Serial.println(" degrees");
    }

    Serial.println("=================================================================================================================================");

  } else {
    Serial.println("캘리브레이션 실패...");
  }

  delete[] real;
  delete[] imag;
}

void impedanceMeasurementCOB() {
  getAddressInput();
  setMuxGroup();
  frequencySweepRaw(startFreq, frequencyUnit, numIncrements);
}

void impedanceMeasurementRcal() {
  frequencySweepRaw(startFreq, frequencyUnit, numIncrements);
}

void showSweepMenu() {
  // 시작 주파수 입력
  while (true) {
    Serial.print("시작 주파수를 입력하세요 (1~100 kHz): ");
    flushSerialBuffer();
    delay(10);
    while (Serial.available() == 0) {
    }
    startFreq = Serial.readStringUntil('\n').toInt();
    if (startFreq >= 1 && startFreq <= 100) {
      startFreq *= 1000; // kHz를 Hz로 변환
      Serial.print("설정된 시작 주파수: ");
      Serial.print(startFreq);
      Serial.println(" Hz");
      break;
    } else {
      Serial.println("잘못된 입력입니다. 1~100 kHz의 값을 입력하세요.");
    }
  }

  // 주파수 증가량 입력
  while (true) {
    Serial.print("주파수 증가량을 입력하세요 (1~1000 Hz): ");
    flushSerialBuffer();
    delay(10);
    while (Serial.available() == 0) {
    }
    frequencyUnit = Serial.readStringUntil('\n').toInt();
    if (frequencyUnit >= 1 && frequencyUnit <= 1000) {
      Serial.print("설정된 주파수 증가량: ");
      Serial.print(frequencyUnit);
      Serial.println(" Hz");
      break;
    } else {
      Serial.println("잘못된 입력입니다. 1~1000 Hz의 값을 입력하세요.");
    }
  }

  // 측정 횟수 입력
  while (true) {
    Serial.print("측정 횟수를 입력하세요 (1~100): ");
    flushSerialBuffer();
    delay(10);
    while (Serial.available() == 0) {
    }
    numIncrements = Serial.readStringUntil('\n').toInt();
    if (numIncrements >= 1 && numIncrements <= 100) {
      Serial.print("설정된 측정 횟수: ");
      Serial.print(numIncrements);
      Serial.println(" 회");
      break;
    } else {
      Serial.println("잘못된 입력입니다. 1~100의 값을 입력하세요.");
    }
  }

  // Settling Time Cycles 입력
  int settlingTimeCycles;
  while (true) {
    Serial.print("Settling Time Cycles를 입력하세요 (0~511): ");
    flushSerialBuffer();
    delay(10);
    while (Serial.available() == 0) {
    }
    settlingTimeCycles = Serial.readStringUntil('\n').toInt();
    if (settlingTimeCycles >= 0 && settlingTimeCycles <= 511) {
      ad5933.setSettlingCycles(settlingTimeCycles);  // 인스턴스 사용(헤더 Settling Cycle 설정 static 미사용으로 인하여 추가)
      Serial.print("설정된 Settling Time Cycles: ");
      Serial.println(settlingTimeCycles);
      break;
    } else {
      Serial.println("잘못된 입력입니다. 0~511 사이의 값을 입력하세요.");
    }
  }

  // Output Excitation Range 설정
  int excitationRange;
  while (true) {
    Serial.print("Output Excitation Range를 선택하세요 (1: 2 Vpp, 2: 1 Vpp, 3: 0.4 Vpp, 4: 0.2 Vpp): ");
    flushSerialBuffer();
    delay(10);
    while (Serial.available() == 0) {
    }
    excitationRange = Serial.readStringUntil('\n').toInt();
    switch (excitationRange) {
      case 1:
        ad5933.setRange(CTRL_OUTPUT_RANGE_1);
        Serial.println("2 Vpp (Range 1)로 설정되었습니다.");
        break;
      case 2:
        ad5933.setRange(CTRL_OUTPUT_RANGE_2);
        Serial.println("1 Vpp (Range 2)로 설정되었습니다.");
        break;
      case 3:
        ad5933.setRange(CTRL_OUTPUT_RANGE_3);
        Serial.println("0.4 Vpp (Range 3)로 설정되었습니다.");
        break;
      case 4:
        ad5933.setRange(CTRL_OUTPUT_RANGE_4);
        Serial.println("0.2 Vpp (Range 4)로 설정되었습니다.");
        break;
      default:
        Serial.println("잘못된 입력입니다. 1~4 중 하나를 선택하세요.");
        continue;
    }
    break;
  }

  // PGA Control 설정
  int pgaControl;
  while (true) {
    Serial.print("PGA Gain을 선택하세요 (1 또는 5): ");
    flushSerialBuffer();
    delay(10);
    while (Serial.available() == 0) {
    }
    pgaControl = Serial.readStringUntil('\n').toInt();
    if (pgaControl == 1 || pgaControl == 5) {
      AD5933::setPGAGain(pgaControl == 1 ? PGA_GAIN_X1 : PGA_GAIN_X5);
      Serial.print("PGA Gain: x");
      Serial.print(pgaControl);
      Serial.println("로 설정완료!");
      break;
    } else {
      Serial.println("잘못된 입력입니다. 1 또는 5를 입력하세요.");
    }
  }

  // Calibration Impedance 입력
  while (true) {
    Serial.print("Calibration Impedance를 입력하세요 (옴 단위, 양의 정수): ");
    flushSerialBuffer();
    delay(10);
    while (Serial.available() == 0) {
    }
    refResist = Serial.readStringUntil('\n').toInt();
    if (refResist > 0) {
      Serial.print("설정된 Calibration Impedance: ");
      Serial.print(refResist);
      Serial.println(" ohm");
      break;
    } else {
      Serial.println("잘못된 입력입니다. 올바른 Calibration Impedance 값을 입력하세요.");
    }
  }

  // AD5933 설정 완료 메시지 출력
  if (!(AD5933::setStartFrequency(startFreq) &&
        AD5933::setIncrementFrequency(frequencyUnit) &&
        AD5933::setNumberIncrements(numIncrements))) {
    Serial.println("주파수 스윕 설정 실패.");
  } else {
    Serial.println("주파수 스윕 설정 완료.");
  }
}

void getAddressInput() {
    int xAddress[7];
    int yAddress[7];
    const int X_AXIS_PINS[7] = {X_AXIS_ADDR_0, X_AXIS_ADDR_1, X_AXIS_ADDR_2, X_AXIS_ADDR_3, X_AXIS_ADDR_4, X_AXIS_ADDR_5, X_AXIS_ADDR_6};
    const int Y_AXIS_PINS[7] = {Y_AXIS_ADDR_0, Y_AXIS_ADDR_1, Y_AXIS_ADDR_2, Y_AXIS_ADDR_3, Y_AXIS_ADDR_4, Y_AXIS_ADDR_5, Y_AXIS_ADDR_6};

    // X축 주소 설정 및 확인
    Serial.println("X축 Address (7자리, 각 비트는 0 / 1로 입력):");
    for (int i = 0; i < 7; i++) {
        while (1) {
            Serial.print("X Axis Address ");
            Serial.print(i);
            Serial.print(" (0 또는 1): ");
            while (Serial.available() == 0) {
                // 대기
            }
            char c = Serial.read();
            flushSerialBuffer();
            if (c == '0' || c == '1') {
                int bitValue = c - '0';
                xAddress[i] = bitValue;
                Serial.print("X축 Address ");
                Serial.print(i);
                Serial.print(" 설정됨: ");
                Serial.println(bitValue);

                // 해당 핀에 HIGH 또는 LOW 설정
                digitalWrite(X_AXIS_PINS[i], bitValue == 1 ? HIGH : LOW);
                delay(10); // 핀 안정화 대기

                int actualValue = digitalRead(X_AXIS_PINS[i]);
                Serial.print("(설정값: ");
                Serial.print(bitValue);
                Serial.print(") - (핀 상태: ");
                Serial.print(actualValue);
                Serial.println(")");
                break;
            } else {
                Serial.println("잘못된 입력. 0 또는 1 입력.");
                flushSerialBuffer();
            }
        }
    }

    // X축 주소를 문자열로 생성
    String xAddrStr = "";
    for(int i = 0; i < 7; i++) {
        xAddrStr += String(xAddress[i]);
    }
    Serial.print("설정된 X축 Address : ");
    Serial.println(xAddrStr);

    // Y축 주소 설정 및 확인
    Serial.println("Y축 Address (7자리, 각 비트는 0 또는 1로 입력):");
    for (int i = 0; i < 7; i++) {
        while (1) {
            Serial.print("Y Axis Address ");
            Serial.print(i);
            Serial.print(" (0 또는 1): ");
            while (Serial.available() == 0) {
                // 대기
            }
            char c = Serial.read();
            flushSerialBuffer();
            if (c == '0' || c == '1') {
                int bitValue = c - '0';
                yAddress[i] = bitValue;
                Serial.print("Y축 Address ");
                Serial.print(i);
                Serial.print(" 설정됨: ");
                Serial.println(bitValue);

                // 해당 핀에 HIGH 또는 LOW 설정
                digitalWrite(Y_AXIS_PINS[i], bitValue == 1 ? HIGH : LOW);
                delay(10); // 핀 안정화 대기

                int actualValue = digitalRead(Y_AXIS_PINS[i]);
                Serial.print("(설정값: ");
                Serial.print(bitValue);
                Serial.print(") - (핀 상태: ");
                Serial.print(actualValue);
                Serial.println(")");
                break;
            } else {
                Serial.println("잘못된 입력. 0 또는 1 입력.");
                flushSerialBuffer();
            }
        }
    }

    // Y축 주소를 문자열로 생성
    String yAddrStr = "";
    for(int i = 0; i < 7; i++) {
        yAddrStr += String(yAddress[i]);
    }
    Serial.print("설정된 Y축 Address : ");
    Serial.println(yAddrStr);

    // 전체 설정 요약 출력
    Serial.print("설정된 X축 Address : ");
    Serial.print(xAddrStr);
    Serial.print(", Y축 Address : ");
    Serial.println(yAddrStr);
}

void setMuxGroup() {
    int group = 0;

    while (1) {
        Serial.print("MUX 그룹을 선택하세요 (1, 2, 3, 4): ");
        flushSerialBuffer();
        delay(10);

        while (Serial.available() == 0) {
            // 대기
        }

        String inputStr = Serial.readStringUntil('\n');
        group = inputStr.toInt();

        if (group >= 1 && group <= 4) {
            Serial.print("그룹 ");
            Serial.print(group);
            Serial.println(" 선택");
            break;
        } else {
            Serial.println("잘못된 입력. 1~4 사이의 값을 입력하세요.");
        }
    }

    // 그룹에 따른 MUX 스위치 설정
    switch (group) {
        case 1: // 00
            digitalWrite(ANALOG_MUX_SWITCH_0, LOW);
            digitalWrite(ANALOG_MUX_SWITCH_1, LOW);
            digitalWrite(DIGITAL_MUX_SWITCH_0, LOW);
            digitalWrite(DIGITAL_MUX_SWITCH_1, LOW);
            break;
        case 2: // 01
            digitalWrite(ANALOG_MUX_SWITCH_0, HIGH);
            digitalWrite(ANALOG_MUX_SWITCH_1, LOW);
            digitalWrite(DIGITAL_MUX_SWITCH_0, HIGH);
            digitalWrite(DIGITAL_MUX_SWITCH_1, LOW);
            break;
        case 3: // 10
            digitalWrite(ANALOG_MUX_SWITCH_0, LOW);
            digitalWrite(ANALOG_MUX_SWITCH_1, HIGH);
            digitalWrite(DIGITAL_MUX_SWITCH_0, LOW);
            digitalWrite(DIGITAL_MUX_SWITCH_1, HIGH);
            break;
        case 4: // 11
            digitalWrite(ANALOG_MUX_SWITCH_0, HIGH);
            digitalWrite(ANALOG_MUX_SWITCH_1, HIGH);
            digitalWrite(DIGITAL_MUX_SWITCH_0, HIGH);
            digitalWrite(DIGITAL_MUX_SWITCH_1, HIGH);
            break;
    }

    Serial.println("MUX 스위치가 설정되었습니다.");
}

void frequencySweepRaw(int startFreq, int frequencyUnit, int numIncrements) {
  int real, imag, i = 0;
  double cfreq = startFreq / 1000.0;

  if (!(AD5933::setPowerMode(POWER_STANDBY) && 
        AD5933::setControlMode(CTRL_INIT_START_FREQ) && 
        AD5933::setControlMode(CTRL_START_FREQ_SWEEP))) {
    Serial.println("Could not initialize frequency sweep...");
    return;
  }

  Serial.println("=================================================================================================================================");

  while ((AD5933::readStatusRegister() & STATUS_SWEEP_DONE) != STATUS_SWEEP_DONE) {
    if (!AD5933::getComplexData(&real, &imag)) {
      Serial.println("Could not get raw frequency data...");
      real = 0;
      imag = 0;
    }

    double magnitude = sqrt(pow(real, 2) + pow(imag, 2));
    double impedance = 1 / (magnitude * gain[i]); // 보정된 임피던스 계산
    double rawPhase = atan2(imag, real) * (180.0 / M_PI); // 측정 위상 (degrees)

    // 측정된 rawPhase를 시스템 위상 방식에 맞추어 보정
    double correctedPhase;
    correctedPhase = rawPhase - phase[i];

    if (correctedPhase < 0.0) {
        correctedPhase += 360.0;
    } else if (correctedPhase >= 360.0) {
        correctedPhase -= 360.0;
    }

    // 보정된 위상을 이용한 실수(Rreal)와 허수(Ximaginary) 성분 계산
    double Rreal = impedance * cos(correctedPhase * M_PI / 180.0); // 실수 성분
    double Ximaginary = impedance * sin(correctedPhase * M_PI / 180.0); // 허수 성분

    // 주파수 및 결과 출력
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
    Serial.println("Could not set to standby...");
  }
}
