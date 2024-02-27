echo "COPIAR TODO EL CONTENIDO DE ESTA CARPETA A C:\Windows\System32 y ejecutar registrar.dll"

regsvr32 comcat.dll
regsvr32 COMDLG32.OCX
regsvr32 CSWSK32.OCX
regsvr32 MSCOMCT2.OCX
regsvr32 MSCOMCTL.OCX
regsvr32 MSFLXGRD.OCX
regsvr32 MSINET.oca
regsvr32 MSINET.OCX
regsvr32 msvbvm50.dll
regsvr32 msvbvm60.dll
regsvr32 MSWINSCK.OCX
regsvr32 oleaut32.dll
regsvr32 olepro32.dll
regsvr32 RICHTX32.OCX

regsvr32 -u riched32.dll
regsvr32 riched32.dll