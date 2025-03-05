I conducted an investigation of a malicious macro embedded in a Microsoft Office document and demonstrated a method for decrypting it using Python.

First, I obtained the file MrPhisher.docm, which contained a macro written in VBA (Visual Basic for Applications). I opened it in LibreOffice Writer, navigated to Tools -> Macros -> Edit Macros, and found the malicious code under the New Macros section.

After analyzing the code, I discovered that it uses an array of numbers, each of which undergoes an XOR operation with its index in a loop. The result is an encoded message that is then displayed on the screen.

Instead of executing the potentially dangerous VBA code, I decided to replicate its logic in Python. I wrote a script that performs the same XOR operation to reconstruct the original message. When I ran the script, I successfully decrypted the text, which turned out to be a flag.

This project helped me gain a deeper understanding of how attackers use macros in phishing attacks and taught me safe methods for analyzing malicious code without executing it.
