import snap7
from snap7.util import *


# Konfiguracja
PLC_IP = '192.168.1.1'  # Adres IP PLC
RACK = 0                # Rack, zazwyczaj 0 dla S7-1500
SLOT = 1                # Slot, zazwyczaj 1 dla CPU

# Inicjalizacja klienta
client = snap7.client.Client()
client.connect(PLC_IP, RACK, SLOT)

if client.get_connected():
    print("Połączono ze sterownikiem S7-1500")
else:
    print("Nie udało się połączyć")

# Zakończenie połączenia
client.disconnect()
print("Rozłączono z PLC")
