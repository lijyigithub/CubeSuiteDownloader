import clr
clr.AddReferenceByPartialName("System")
clr.AddReferenceByPartialName("System.Windows.Forms")
clr.AddReferenceByPartialName("System.Drawing")
clr.AddReferenceByPartialName("IronPython")
clr.AddReferenceByPartialName("Microsoft.Scripting")
import System
import os
from System.Windows.Forms import *
from System.Drawing import *
from System.Threading import Thread
from System.Threading import ThreadStart
from System.Threading import AutoResetEvent
from System.ComponentModel import BackgroundWorker
import sys
import time
import tempfile

filepath = ""
rst = RichTextBox()
fopen = Button()
fwrite = Button()
l_mtime = 0
l_ctime = 0
t = BackgroundWorker()

mcolor = {
    "INFO":Color.Black,
    "OK":Color.Green,
    "ERR":Color.Red
}

def wker(a, b):
    tk = Timer()
    tk.Tick += lambda m, n:Application.DoEvents()
    tk.Interval = 100
    tk.Enabled = True


def log(msg, c="INFO"):
    outs = time.strftime("[%Y-%m-%d %H:%M:%S] ", time.localtime())
    # rst.Font = Font("Consolas", 9)
    # rst.Text = rst.Text + "\r\n" + outs
    rst.SelectionStart = rst.Text.Length
    rst.SelectionColor = mcolor["INFO"]
    rst.SelectedText = "\r\n%s" % outs
    rst.SelectionStart = rst.Text.Length
    rst.SelectionLength = 0
    rst.SelectionColor = mcolor[c]
    rst.SelectedText =  msg
    rst.SelectionStart = rst.Text.Length
    rst.SelectionLength = 0
    rst.ScrollToCaret()

def sfile(obj, evt):
    global filepath, l_mtime, l_ctime
    f = OpenFileDialog()
    f.InitialDirectory = "d:\\"
    f.Filter = "hex files (*.hex)|*.hex|All files (*.*)|*.*"
    f.RestoreDirectory = True
    f.ShowHelp = True  # shen me gui
    if f.ShowDialog()==DialogResult.OK:
        log("Select File '%s'" % f.FileName)
        log("File Create Time: %s" % time.strftime("[%Y-%m-%d %H:%M:%S] ", time.gmtime(os.stat(f.FileName).st_ctime)))
        log("File Modify Time: %s" % time.strftime("[%Y-%m-%d %H:%M:%S] ", time.gmtime(os.stat(f.FileName).st_mtime)))
        l_mtime = os.stat(f.FileName).st_mtime
        l_ctime = os.stat(f.FileName).st_ctime
        filepath = f.FileName

def dfile(a, b):
    if not os.path.exists(filepath):
        log("File Not Found: '%s'" % filepath, "ERR")
        return
    log("=================================================")
    if (l_mtime != os.stat(filepath).st_mtime or
        l_ctime != os.stat(filepath).st_ctime):
        log("File Changed")
        log("File Create Time: %s" % time.strftime("[%Y-%m-%d %H:%M:%S] ", time.gmtime(os.stat(filepath).st_ctime)))
        log("File Modify Time: %s" % time.strftime("[%Y-%m-%d %H:%M:%S] ", time.gmtime(os.stat(filepath).st_mtime)))
    # bts = open(filepath, 'rb').read()
    # file_sum = 0
    # for b in bts:
    #     file_sum += ord(b)
    # log("File CheckSum %08X" % file_sum)
    log("Begin download")
    log("Connect to MINICUBE II")
    ret = None
    try:
        ret = debugger.Connect()
    except:
        log(sys.exc_value.__str__(), "ERR")
    if ret:
        log("Connect Sucessed", "OK")
    else:
        log("Connect Faild", "ERR")
        return

    log("Erase Chip")
    try:
        ret = debugger.Erase()
    except:
        log(sys.exc_value.__str__(), "ERR")
    if ret:
        log("Erase Sucessed", "OK")
    else:
        log("Erase Faild", "ERR")
        return

    log("Download HEX File")
    try:
        ret = debugger.Download.Hex(filepath)
    except:
        log(sys.exc_value.__str__(), "ERR")
    if ret:
        log("Download Sucessed", "OK")
    else:
        log("Download Faild", "ERR")
        return

    log("Disconnect with MINICUBE II")
    debugger.Disconnect()

def begin_dfile(a, b):
    def apply(bo):
        fopen.Enabled = bo
        fwrite.Enabled = bo
    apply(False)
    t = BackgroundWorker()
    t.DoWork += dfile
    t.RunWorkerCompleted += lambda a, b:apply(True)
    t.RunWorkerAsync()

dispatcher = Form()
dispatcher.Width = 800
dispatcher.Height = 600
dispatcher.MaximizeBox = False
# dispatcher.TopMost = True
dispatcher.Text = "NEC HEX Downloader"

rst.Dock = DockStyle.Top
rst.Height = 600 - 120
rst.Multiline = True
rst.ReadOnly = True
rst.ScrollBars = RichTextBoxScrollBars.Vertical
rst.BackColor = Color.White
# rst.Font = Font("Consolas", 9)
rst.DetectUrls = False
dispatcher.Controls.Add(rst)

fopen.Text = r"&Open Hex File"
fopen.Click += sfile
fopen.Height = 48
fopen.Width = 120
fopen.Top = 490
fopen.Left = 40
dispatcher.Controls.Add(fopen)

fwrite.Text = r"&Download"
fwrite.Click += begin_dfile
fwrite.Height = 48
fwrite.Width = 120
fwrite.Top = 490
fwrite.Left = 600
dispatcher.Controls.Add(fwrite)

dispatcher.StartPosition = FormStartPosition.CenterScreen
log("Program Begin")
Application.Run(dispatcher)
del t
