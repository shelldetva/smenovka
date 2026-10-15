' Spustí sync.py úplne skryte (bez okna). Výpis ukladá do %TEMP%\smenovka_log.txt
Set s = CreateObject("WScript.Shell")
s.CurrentDirectory = "C:\!SHELLAutomatizacia\SmenovkaWEB"
s.Run "cmd /c python -u sync.py > ""%TEMP%\smenovka_log.txt"" 2>&1", 0, False
