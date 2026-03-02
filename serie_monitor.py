import serial

SERIAL_PORT = "/dev/ttyUSB0"   # axusta ao teu porto: /dev/ttyACM0, COM3, etc.
BAUD_RATE   = 9600

def main():
    print(f"Conectando a {SERIAL_PORT} a {BAUD_RATE} baud... (Ctrl+C para saír)")
    with serial.Serial(SERIAL_PORT, BAUD_RATE, timeout=2) as ser:
        while True:
            liña = ser.readline()
            if liña:
                print(liña.decode("utf-8", errors="replace"), end="")

if __name__ == "__main__":
    main()
