# Office365
Activate Office 365

_**MANNUAL ACTIVATION**_
1. Open command prompt as admin.
2. Navigate to your Office folder. If you install your Office in the ProgramFiles folder, the path will be “%ProgramFiles%\Microsoft Office\Office16” or “%ProgramFiles(x86)%\Microsoft Office\Office16.
3. Convert your Office license to volume one if possible. If your Office is got from Microsoft, this step is required. On the contrary, if you install Office from a Volume ISO file, this is optional so just skip it if you want.
4. Use KMS client key to activate your Office. Make sure your PC is connected to the internet.
5. DONE. Enjoy using MS office 365
6. ALL CODES ARE PROVIDED SEPERATELY FOR EASE.

------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

CODES :

**Navigate and locate Office 365**

cd /d %ProgramFiles%\Microsoft Office\Office16
OR 
cd /d %ProgramFiles(x86)%\Microsoft Office\Office16

**Convert to Volume**

for /f %x in ('dir /b ..\root\Licenses16\proplusvl_kms*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%x"

**Install Activation key**
   
 cscript ospp.vbs /inpkey:XQNVK-8JYDB-WJ9W3-YJ8YR-WFG99
 
 cscript ospp.vbs /unpkey:BTDRB >nul
 
 cscript ospp.vbs /unpkey:KHGM9 >nul
 
 cscript ospp.vbs /unpkey:CPQVG >nul
 
 cscript ospp.vbs /sethst:107.175.77.7
 
 cscript ospp.vbs /setprt:1688
 
 cscript ospp.vbs /act

-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

_**AUTO ACTIVATION using BATCH SCRIPT**_
1. Copy the code in "Manual Acttivation" into a new text document or Download Office.cmd
2. Save it as a batch file. (eg. office365.cmd)
3. Run the batch file with admin rights. (important!)
   diable any active antivirus


ENJOY USING OFFICE 365
