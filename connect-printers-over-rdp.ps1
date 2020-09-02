#---------------------------------------------------------------------------#
#                                                                           #
#   Damit das Script korrekt funktioniert muessen beide Drucker lokal       #
#   eingrichtet sein und die Weiterleitung von Druckern in der              #
#   RDP-Verbindungsdatei muss aktiviert sein.                               #
#                                                                           #
#---------------------------------------------------------------------------#

  

# Hier das Netz im Büro eintragen
$officeip = '192.168.120.0'

# Hier das Netz im Home-Office eintragen
$homeip = '192.168.1.0' 

# Hier den Drucker im Büro eintragen
$officeprinter = 'DRUCKER BUERO'

# Hier den Drucker im Home-Office eintragen
$homeprinter = 'DRUCKER HOME'

# Hier die Remotedesktop-Verbindungs-Datei angeben
$rdp = 'C:\Users\Public\Desktop\WTS1.rdp'

 
#                        Ab hier nichts mehr bearbeiten!                    #
# ------------------------------------------------------------------------- #


# Variablen anpassen
$officeip = $officeip.TrimEnd("0")
$officeip = $officeip+'*'
$homeip = $homeip.TrimEnd("0")
$homeip = $homeip+'*'

# IP-Adresse raussuchen
$ip = (
    Get-NetIPConfiguration |
    Where-Object {
        $_.IPv4DefaultGateway -ne $null -and
        $_.NetAdapter.Status -ne "Disconnected"
    }
).IPv4Address.IPAddress

# Drucker auswählen
if($ip -like "$officeip"){
	echo "Standort: Office. Setzte Drucker auf $officeprinter."
    rundll32 printui.dll,PrintUIEntry /y /n "$officeprinter"
    }
elseif($ip -like "$homeip"){
    echo "Standort: Home-Office. Setzte Drucker auf $homeprinter."
    rundll32 printui.dll,PrintUIEntry /y /n $homeprinter
    }
else{
    echo "Unbekannter Standort"
    }

# Terminalserver-Sitzung starten
mstsc $rdp
