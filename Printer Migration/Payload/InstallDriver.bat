:: Install Sharp Driver
certutil -addstore "TrustedPublisher" Sharpcert.cer
rundll32 printui.dll,PrintUIEntry /ia /f "\\Iwmdocs\iwm\CIWMB-INFOTECH\Network\Printers\SHARP-MX-4141N\SOFTWARE-CDs\CD1\Drivers\Printer\English\PS\64bit\ss0hmenu.inf" /m "SHARP MX-4141N PS"
certutil -delstore "TrustedPublisher" Shapercer.cer