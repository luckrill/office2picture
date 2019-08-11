#!/usr/bin/env python
# -*- coding: utf-8 -*-
import wx
import os
import sys
import subprocess
import codecs
import shutil
import time
# import win32com.client as win32
import subprocess
import importlib
import smtplib
from email.message import EmailMessage

# importlib.reload(sys)
# sys.setdefaultencoding("utf-8")

L = wx.Locale()
_ = wx.GetTranslation

def anyTrue(predicate, sequence):
    return True in list(map(predicate, sequence))

pdf_exts = [".pdf", ".PDF"]
word_exts = [".docx", ".doc", ".DOCX", ".DOC"]
ppt_exts = [".pptx", ".ppt", ".PPTX", ".PPT"]
excel_exts = [".xlsx", ".xls", ".XLSX", ".XLS"]
office_exts = [".docx", ".doc", ".pptx", ".ppt", ".xlsx", ".xls", ".DOCX", ".DOC", ".PPTX", ".PPT", ".XLSX", ".XLS", ".pdf", ".PDF"]

class DBfilenames():
    filenames = []
    basepath = ""
    config_file = ""
    about_file = ""
    output_path = "d:\office_output"
    current = {"convert_to":"Picture", "lang":"English", "image_quality":"Good", "image_format":"png"}
    gs = "tools/gswin64.exe"

    def __init__(self):
        self.basepath = self.check_path()
        self.config_file = os.path.join(self.basepath, "doc", "Config")
        self.about_file = os.path.join(self.basepath, "doc", "About")

    def check_path(self):
        basepath = os.path.abspath(os.path.dirname(sys.argv[0]))
        config_file = os.path.join(basepath, "doc", "Config")
        try:
            open(config_file, 'a+')
        except IOError as e:
            app_path = os.environ['PROGRAMDATA']
            office2picture = os.path.join(app_path, "office2picture")
            if not os.path.exists(office2picture):
                os.mkdir(office2picture)
            if not os.path.exists(office2picture+os.sep+"doc"):
                shutil.copytree(basepath+os.sep+"doc", office2picture+os.sep+"doc")
            basepath = office2picture

        return basepath

    def ConfigLoad(self):
        """load config info into current dict from config file"""
        if not os.path.exists(dbfilenames.config_file):
            return
        fd = open(dbfilenames.config_file, mode='r', encoding="utf-8")
        while True:
            line = fd.readline()
            # print line
            if not line:
                break
            line = line.strip()
            if (line.startswith("#")):
                continue
            li = line.split("=")
            li = [x.strip() for x in li]
            if li[0] == "convert_to":
                dbfilenames.current["convert_to"] = li[1]
            elif li[0] == "lang":
                dbfilenames.current["lang"] = li[1]
            elif li[0] == "image_quality":
                dbfilenames.current["image_quality"] = li[1]
            elif li[0] == "image_format":
                dbfilenames.current["image_format"] = li[1]

        fd.close()

    def ConfigSave(self):
        """save window config info to current dict and config file"""
        fd = open(dbfilenames.config_file, 'w', encoding="utf-8")
        for key in list(dbfilenames.current.keys()):
            fd.write(str(key) + " = " + str(dbfilenames.current[key]) + "\n")
        fd.close()

class MainFrame(wx.Frame):
    def __init__(self):
        wx.Frame.__init__(self, None, -1, "Office2Picture", size=(900, 800))
        self.SetIcon(wx.Icon('myicon.ico', wx.BITMAP_TYPE_ICO))

        self.Centre()
        panel = wx.Panel(self)
        # self.CreateStatusBar()

        self.input_path = ""
        self.input_path_len = 0
        self.setup_frame = 0
        self.first_start = True
        self.first_opendir = True

        font = wx.SystemSettings.GetFont(wx.SYS_DEFAULT_GUI_FONT)
        font.SetPointSize(font.GetPointSize() + 2)
        panel.SetFont(font)

        button_output = wx.Button(panel, -1, _("Output"), size=(70, -1))
        self.text_output = wx.TextCtrl(panel, -1, "d:\office_output", size=(300, -1), style=wx.TE_PROCESS_ENTER|wx.TE_LEFT)
        self.button_to_picture = wx.Button(panel, -1, label=_("Change to Picture"))
        button_setup = wx.Button(panel, -1, label=_("Setup"))
        button_tellme = wx.Button(panel, -1, label=_("TellMe"))
        button_tellme.SetToolTip("TellMe Something ... Please")
        button_close = wx.Button(panel, -1, label=_("Close"))

        label_listbox = wx.StaticText(panel, -1, _("Select Files"), size=(100, -1))
        label_logs = wx.StaticText(panel, -1, _("Work Logs"), size=(100, -1))
        self.listbox = wx.ListBox(panel, -1, (20, 20), (380, 500), [], wx.LB_EXTENDED)
        self.text_multi_text = wx.TextCtrl(panel, -1, "", size=(380, 500), style=wx.TE_MULTILINE)
        self.text_multi_text.SetValue("Enjoy Light \nPlease add office files at first \nSupport Word, PowerPoint, Excel, PDF document \nThen execute 'Change to Picture'")

        self.list_add_files = wx.Button(panel, -1, label=_("Add Files"))
        list_add_path = wx.Button(panel, -1, label=_("Add Path"))
        list_clear = wx.Button(panel, -1, label=_("Clear"))
        list_remove = wx.Button(panel, -1, label=_("Remove"))


        sizer = wx.BoxSizer(wx.VERTICAL)
        vbox_list = wx.BoxSizer(wx.HORIZONTAL)
        vbox_list_cmd = wx.BoxSizer(wx.HORIZONTAL)


        vbox_output = wx.BoxSizer(wx.HORIZONTAL)
        vbox_output.Add((10, 10))
        vbox_output.Add(button_output, 0, wx.ALIGN_LEFT|wx.ALIGN_CENTER)
        vbox_output.Add((5, 5))
        vbox_output.Add(self.text_output, 0, wx.ALIGN_LEFT|wx.ALIGN_CENTER)
        vbox_output.Add((50, 50))
        vbox_output.Add(self.button_to_picture, 0, wx.ALIGN_CENTER)
        vbox_output.Add((10, 10))
        vbox_output.Add(button_setup, 0, wx.ALIGN_CENTER)
        vbox_output.Add((10, 10))
        vbox_output.Add(button_tellme, 0, wx.ALIGN_CENTER)
        vbox_output.Add((10, 10))
        vbox_output.Add(button_close, 0, wx.ALIGN_CENTER)

        vbox_listbox = wx.BoxSizer(wx.VERTICAL)
        vbox_listbox.Add(label_listbox, 0, wx.ALIGN_LEFT)
        vbox_listbox.Add(self.listbox, 1, wx.EXPAND)

        vbox_multi_text = wx.BoxSizer(wx.VERTICAL)
        vbox_multi_text.Add(label_logs, 0, wx.ALIGN_LEFT)
        vbox_multi_text.Add(self.text_multi_text, 1, wx.EXPAND)

        vbox_list_cmd.Add(self.list_add_files, 0, wx.ALIGN_CENTER)
        vbox_list_cmd.Add((8, 8))
        vbox_list_cmd.Add(list_add_path, 0, wx.ALIGN_CENTER)
        vbox_list_cmd.Add((20, 20))
        vbox_list_cmd.Add(list_clear, 0, wx.ALIGN_CENTER)
        vbox_list_cmd.Add((8, 8))
        vbox_list_cmd.Add(list_remove, 0, wx.ALIGN_CENTER)


        vbox_list.Add((10, 10))
        vbox_list.Add(vbox_listbox, 1, wx.EXPAND)
        vbox_list.Add((10, 10))

        vbox_list.Add((10, 10))
        vbox_list.Add(vbox_multi_text, 1, wx.EXPAND)
        vbox_list.Add((10, 10))

        sizer.Add((5, 5))
        sizer.Add(vbox_output, flag=wx.ALIGN_LEFT)
        sizer.Add(vbox_list_cmd, 0, wx.ALIGN_LEFT)
        sizer.Add((10, 10))
        sizer.Add(vbox_list, 1, wx.EXPAND)

        panel.SetSizer(sizer)
        panel.Layout()

        self.listbox.SetFocus()
        # self.list_add_files.SetFocus()
        self.list_add_files.SetDefault()


        self.Bind(wx.EVT_BUTTON, self.OnAddFiles, self.list_add_files)
        self.Bind(wx.EVT_BUTTON, self.OnAddPath, list_add_path)
        self.Bind(wx.EVT_BUTTON, self.OnListClear, list_clear)
        self.Bind(wx.EVT_BUTTON, self.OnListRemove, list_remove)


        self.Bind(wx.EVT_BUTTON, self.OnToPicture, self.button_to_picture)
        self.Bind(wx.EVT_BUTTON, self.OnOutput, button_output)
        self.Bind(wx.EVT_BUTTON, self.OnSetup, button_setup)
        self.Bind(wx.EVT_BUTTON, self.OnTellMe, button_tellme)
        self.Bind(wx.EVT_BUTTON, self.OnClose, button_close)
        self.Bind(wx.EVT_CLOSE, self.OnExit)

    def OnOutput(self, event):
        input_path = ""
        dialog = wx.DirDialog(None,"Choose a directory:",
                              style=wx.DD_DEFAULT_STYLE|wx.DD_DIR_MUST_EXIST)

        if dialog.ShowModal() == wx.ID_OK:
            input_path = dialog.GetPath()
            self.text_output.SetValue(input_path)

        dialog.Destroy()

    def OnAddFiles(self, event):
        dialog = wx.FileDialog(self, "Choose Office Files", "", "",
                                "*.*",
                                style = wx.FD_OPEN | wx.FD_FILE_MUST_EXIST | wx.FD_MULTIPLE)
        if dialog.ShowModal() == wx.ID_OK:
            for name in dialog.GetFilenames():
                if anyTrue(name.endswith, office_exts):
                    path = dialog.GetDirectory()
                    dbfilenames.filenames.append([path, name])
        dialog.Destroy()

        self.ListBoxDataUpdate()
        self.button_to_picture.SetFocus()
        self.button_to_picture.SetDefault()

    def OnAddPath(self, event):
        input_path = ""
        dialog = wx.DirDialog(None,"Choose a directory:",
                              style=wx.DD_DEFAULT_STYLE|wx.DD_DIR_MUST_EXIST)

        if dialog.ShowModal() == wx.ID_OK:
            self.input_path = dialog.GetPath()
            self.input_path_len = len(self.input_path) + 1 # add os.sep

        dialog.Destroy()
        self.AddImageFiles(self.input_path)

        self.ListBoxDataUpdate()
        self.button_to_picture.SetFocus()

        # for name in dbfilenames.filenames:
        #     print name

    def AddImageFiles(self, input_path):
        # 数据重复，容许，不提醒用户，用户自己判断
        if os.path.isdir(input_path):
            os.path.walk(input_path, self.scanfile, 0)

        self.list_num = len(dbfilenames.filenames)

    def scanfile(self, arg, dirname, names):
        for tmpfile in names:
            filename = dirname + os.sep + tmpfile
            if os.path.isfile(filename) and (os.stat(filename)[6] != 0):
                if anyTrue(filename.endswith, office_exts):
                    short_filename = filename[self.input_path_len:]
                    # print dirname, "|", short_filename, "|", tmpfile
                    dbfilenames.filenames.append([self.input_path, short_filename])

    def ListBoxDataUpdate(self):
        if self.first_start == True:
            self.text_multi_text.Clear()
            self.first_start = False

        self.listbox.Clear()
        for li in dbfilenames.filenames:
            self.listbox.Append(li[1])
        if len(dbfilenames.filenames) > 0:
            self.listbox.SetSelection(0)

    def PptToPicture(self, path, filename):
        in_filename = os.path.join(path, filename)
        fileformat = 18
        pos = filename.find(".")
        filename = filename[:pos].strip()
        if dbfilenames.current["convert_to"] == "Picture":
            fileformat = 18
            output_path = os.path.join(dbfilenames.output_path, filename)
            if not os.path.isdir(output_path):
                os.makedirs(output_path)
        elif dbfilenames.current["convert_to"] == "PDF":
            fileformat = 32
            output_path = os.path.join(dbfilenames.output_path, filename + ".pdf")
            filepath = os.path.dirname(output_path)
            if not os.path.exists(filepath):
                os.makedirs(filepath)

        App = win32.DispatchEx("PowerPoint.Application")
        App.Visible = True
        self.text_multi_text.AppendText(_("PowerPoint open: ") + in_filename + "\n")
        presentation = App.Presentations.Open(in_filename, ReadOnly = True)
        self.text_multi_text.AppendText(_("PowerPoint save as png: ") + output_path + "\n")
        presentation.SaveAs(output_path, fileformat)
        presentation.Close()
        App.Quit()
        self.text_multi_text.AppendText(_("PowerPoint save as png done\n\n"))

    def WordToPicture(self, path, filename):
        in_filename = os.path.join(path, filename)

        pos = filename.find(".")
        filename = filename[:pos].strip()
        if dbfilenames.current["convert_to"] == "Picture":
            print("save as picture")
            output_path = os.path.join(dbfilenames.output_path, filename)
            pdftmpfile = os.path.join(dbfilenames.output_path, "tmpfile.pdf")#.encode("utf-8")
            if not os.path.isdir(output_path):
                os.makedirs(output_path)
        elif dbfilenames.current["convert_to"] == "PDF":
            print("save as pdf")
            output_path = os.path.join(dbfilenames.output_path, filename + ".pdf")
            pdftmpfile = output_path
            filepath = os.path.dirname(output_path)
            if not os.path.exists(filepath):
                os.makedirs(filepath)

        App = win32.DispatchEx('Word.Application')
        App.Visible = True
        self.text_multi_text.AppendText(_("Word open: ") + in_filename + "\n")
        presentation = App.Documents.Open(in_filename, ReadOnly = True)
        print(pdftmpfile)
        self.text_multi_text.AppendText(_("Word save as Pdf: ") + pdftmpfile + "\n")
        # exe meet error: win32.constants.wdFormatPDF
        # App.DisplayAlerts = False
        presentation.SaveAs(pdftmpfile, FileFormat=17)
        # App.DisplayAlerts = True
        presentation.Close()
        App.Quit()
        self.text_multi_text.AppendText(_("Word save as Pdf done\n"))
        self.TmpPdfToPicture(pdftmpfile, output_path)
        # self.PdfToPicture(pdftmpfile, output_path)

    def ExcelToPicture(self, path, filename):
        in_filename = os.path.join(path, filename)

        pos = filename.find(".")
        filename = filename[:pos].strip()
        if dbfilenames.current["convert_to"] == "Picture":
            print("picture type")
            output_path = os.path.join(dbfilenames.output_path, filename)
            pdftmpfile = os.path.join(dbfilenames.output_path, "tmpfile.pdf")
            if not os.path.isdir(output_path):
                os.makedirs(output_path)
        elif dbfilenames.current["convert_to"] == "PDF":
            print("pdf type")
            output_path = os.path.join(dbfilenames.output_path, filename + ".pdf")
            pdftmpfile = output_path
            filepath = os.path.dirname(output_path)
            if not os.path.exists(filepath):
                os.makedirs(filepath)

        App = win32.DispatchEx('Excel.Application')
        App.Visible = True

        self.text_multi_text.AppendText(_("Excel open: ") + in_filename + "\n")
        # print in_filename
        presentation = App.Workbooks.Open(in_filename, ReadOnly = True)
        # Excel pdf format = 57
        self.text_multi_text.AppendText(_("Excel save as pdf: ") + pdftmpfile + "\n")
        App.displayalerts = False
        # print pdftmpfile
        presentation.SaveAs(pdftmpfile, FileFormat=57)
        # presentation.SaveAs(pdftmpfile, 57)
        # presentation.SaveAs("pdftmpfile.pdf", 57)
        presentation.Close()
        App.displayalerts = True
        App.Quit()
        self.text_multi_text.AppendText(_("Excel save as Pdf done\n"))
        self.TmpPdfToPicture(pdftmpfile, output_path)
        # self.PdfToPicture(pdftmpfile, output_path)

    def TmpPdfToPicture(self, tmpfile, output_path):
        if dbfilenames.current["convert_to"] == "PDF":
            return

        self.text_multi_text.AppendText(_("Pdf to png: ") + output_path + "\n")

        image_quality = "-r120"
        image_format = "png16m"

        if dbfilenames.current["image_format"] == "png":
            image_format = "png16m"
            output = "-sOutputFile=" + output_path + os.sep + "page-%02d.png"
        elif dbfilenames.current["image_format"] == "jpg":
            image_format = "jpeg"
            output = "-sOutputFile=" + output_path + os.sep + "page-%02d.jpg"
        if dbfilenames.current["image_quality"] == "Good":
            image_quality = "-r120"
        elif dbfilenames.current["image_quality"] == "Better":
            image_quality = "-r200"
        elif dbfilenames.current["image_quality"] == "Best":
            image_quality = "-r300"

        inputfile = os.path.join(dbfilenames.output_path, "tmpfile.pdf")
        subprocess.call([dbfilenames.gs, "-dSAFER", "-dBATCH", "-dNOPAUSE", image_quality, "-sDEVICE=" + image_format, "-dTextAlphaBits=4", output, inputfile])
        self.text_multi_text.AppendText(_("Pdf to png done\n\n"))

    def PdfToPicture(self, path, filename):
        if dbfilenames.current["convert_to"] == "PDF":
            return

        inputfile = os.path.join(path, filename)

        pos = filename.find(".")
        filename = filename[:pos]
        output_path = os.path.join(dbfilenames.output_path, filename)
        if not os.path.isdir(output_path):
            os.makedirs(output_path)

        image_quality = "-r120"
        image_format = "png16m"

        if dbfilenames.current["image_format"] == "png":
            image_format = "png16m"
            output = "-sOutputFile=" + output_path + os.sep + "page-%02d.png"
        elif dbfilenames.current["image_format"] == "jpg":
            image_format = "jpeg"
            output = "-sOutputFile=" + output_path + os.sep + "page-%02d.jpg"
        if dbfilenames.current["image_quality"] == "Good":
            image_quality = "-r120"
        elif dbfilenames.current["image_quality"] == "Better":
            image_quality = "-r200"
        elif dbfilenames.current["image_quality"] == "Best":
            image_quality = "-r300"

        self.text_multi_text.AppendText(_("Pdf to png: ") + output_path + "\n")
        subprocess.call([dbfilenames.gs, "-dSAFER", "-dBATCH", "-dNOPAUSE", "-r150", "-sDEVICE=png16m", "-dTextAlphaBits=4", output, inputfile])
        self.text_multi_text.AppendText(_("Pdf to png done\n\n"))

    def OnToPicture(self, event):
        if self.listbox.GetCount() < 1:
            return

        dbfilenames.output_path = self.text_output.GetValue().strip()
        curtime = time.strftime('%Y-%m-%d-%H-%M-%S',time.localtime(time.time()))
        self.text_multi_text.AppendText(curtime + "\n")
        self.text_multi_text.AppendText(_("Start execute to picture ...... \n\n"))
        index = 0
        for li in dbfilenames.filenames:
            name = li[1]
            self.text_multi_text.AppendText(_("To Picture: ") + name + "\n")
            #self.listbox.DeselectAll()
            self.listbox.SetSelection(index)
            index += 1
            if anyTrue(name.endswith, word_exts):
                self.WordToPicture(li[0], name)
            elif anyTrue(name.endswith, excel_exts):
                self.ExcelToPicture(li[0], name)
            elif anyTrue(name.endswith, ppt_exts):
                self.PptToPicture(li[0], name)
            elif anyTrue(name.endswith, pdf_exts):
                if dbfilenames.current["convert_to"] == "Picture":
                    self.PdfToPicture(li[0], name)
                else:
                    self.text_multi_text.AppendText(_("Skip File: ") + name + "\n")
            else:
                print("error file format")

        #self.listbox.DeselectAll()
        self.listbox.SetSelection(0)
        if self.first_opendir == True:
            subprocess.call(["explorer", self.text_output.GetValue()])
            self.first_opendir = False

    def OnSetup(self, event):
        self.setup_frame = SetupFrame()
        self.setup_frame.Show(True)

    def OnTellMe(self, event):
        dlg = wx.TextEntryDialog(None, "TellMe Something:", caption='TellMe Please', style=wx.TE_MULTILINE | wx.OK | wx.CANCEL | wx.CENTRE)
        if dlg.ShowModal() == wx.ID_OK:
            message = dlg.GetValue()
            print(message)
            self.send_email(message)
        dlg.Destroy()

    def OnClose(self, event):
        # print "on close"
        if self.setup_frame:
            self.setup_frame.Destroy()

        self.Close()
        # self.Destroy()
        event.Skip()

    def OnListClose(self, event):
        self.OnClose(event)

    def OnListClear(self, event):
        if self.listbox.GetCount() < 1:
            return

        self.listbox.Clear()
        del dbfilenames.filenames[:]

    def OnListRemove(self, event):
        totals = self.listbox.GetCount()
        if totals < 1:
            return

        selects = self.listbox.GetSelections()
        for index in selects:
            self.listbox.Delete(index)
            del dbfilenames.filenames[index]
            totals -= 1
            if (index == totals):
                index -= 1
            self.listbox.SetSelection(index)

    def OnListBox(self, event):
        # print "click on listbox"
        pass
        # selects = self.listbox.GetSelections()

    def OnDclickListBox(self, event):
        # print "OnDclickListBox and start imageframe"
        pass

    def OnExit(self, event):
        # self.Close()
        self.Destroy()
        event.Skip()

    def send_email(self, message):
        msg = EmailMessage()
        msg.set_content(message)

        # me == the sender's email address
        # you == the recipient's email address
        msg['Subject'] = 'Office2Picture tellme'
        msg['From'] = "luckrill@163.com"
        msg['To'] = "luckrill@163.com"

        # Send the message via our own SMTP server.
        s = smtplib.SMTP_SSL("smtp.163.com", 465)
        s.login("luckrill", "jiangzx123456")
        s.send_message(msg)
        s.quit()
        pass

class SetupFrame(wx.Frame):
    """System Setup Frame class, sub window."""
    def __init__(self):
        """Create a Frame instance"""
        wx.Frame.__init__(self, None, -1, title=_("System Setup"), size=(750, 580))
        self.Centre()

        panel = wx.Panel(self)
        font = wx.SystemSettings.GetFont(wx.SYS_DEFAULT_GUI_FONT)
        font.SetPointSize(font.GetPointSize() + 2)
        panel.SetFont(font)

        label_lang = wx.StaticText(panel, -1, _("Language"))
        self.radio_en = wx.RadioButton(panel, -1, "English", style=wx.RB_GROUP)
        self.Bind(wx.EVT_RADIOBUTTON, self.OnRadioLanguage, self.radio_en)
        self.radio_zh = wx.RadioButton(panel, -1, "中文")
        self.Bind(wx.EVT_RADIOBUTTON, self.OnRadioLanguage, self.radio_zh)

        label_format = wx.StaticText(panel, -1, _("Format"))
        self.radio_png = wx.RadioButton(panel, -1, "png", style=wx.RB_GROUP)
        self.Bind(wx.EVT_RADIOBUTTON, self.OnRadioFormat, self.radio_png)
        self.radio_jpg = wx.RadioButton(panel, -1, "jpg")
        self.Bind(wx.EVT_RADIOBUTTON, self.OnRadioFormat, self.radio_jpg)

        label_convert_to = wx.StaticText(panel, -1, _("ConvertTo"))
        self.radio_convert_picture = wx.RadioButton(panel, -1, "Picture", style=wx.RB_GROUP)
        self.Bind(wx.EVT_RADIOBUTTON, self.OnRadioConvertTo, self.radio_convert_picture)
        self.radio_convert_pdf = wx.RadioButton(panel, -1, "PDF")
        self.Bind(wx.EVT_RADIOBUTTON, self.OnRadioConvertTo, self.radio_convert_pdf)

        label_quality = wx.StaticText(panel, -1, _("Quality"))
        self.radio_good = wx.RadioButton(panel, -1, "Good", style=wx.RB_GROUP)
        self.Bind(wx.EVT_RADIOBUTTON, self.OnRadioQuality, self.radio_good)
        self.radio_better = wx.RadioButton(panel, -1, "Better")
        self.Bind(wx.EVT_RADIOBUTTON, self.OnRadioQuality, self.radio_better)
        self.radio_best = wx.RadioButton(panel, -1, "Best")
        self.Bind(wx.EVT_RADIOBUTTON, self.OnRadioQuality, self.radio_best)

        label_about = wx.StaticText(panel, -1, _("About"))
        self.text_multi_text = wx.TextCtrl(panel, -1, "", size=(380, 320), style=wx.TE_MULTILINE)
        button_close = wx.Button(panel, -1, label=_("Close"))

        vbox = wx.BoxSizer(wx.VERTICAL)

        vbox_grid = wx.GridBagSizer(hgap=10, vgap=10)
        vbox_grid.Add(label_convert_to, pos=(0,0), flag=wx.ALIGN_RIGHT, border=5)
        vbox_grid.Add(self.radio_convert_picture, pos=(0,1), flag=wx.ALIGN_LEFT, border=5)
        vbox_grid.Add(self.radio_convert_pdf, pos=(0,2), flag=wx.ALIGN_LEFT, border=5)

        vbox_grid.Add(label_format, pos=(1,0), flag=wx.ALIGN_RIGHT, border=5)
        vbox_grid.Add(self.radio_png, pos=(1,1), flag=wx.ALIGN_LEFT, border=5)
        vbox_grid.Add(self.radio_jpg, pos=(1,2), flag=wx.ALIGN_LEFT, border=5)

        vbox_grid.Add(label_quality, pos=(2,0), flag=wx.ALIGN_RIGHT, border=5)
        vbox_grid.Add(self.radio_good, pos=(2,1), flag=wx.ALIGN_LEFT, border=5)
        vbox_grid.Add(self.radio_better, pos=(2,2), flag=wx.ALIGN_LEFT, border=5)
        vbox_grid.Add(self.radio_best, pos=(2,3), flag=wx.ALIGN_LEFT, border=5)

        vbox_grid.Add(label_lang, pos=(3,0), flag=wx.ALIGN_RIGHT, border=5)
        vbox_grid.Add(self.radio_en, pos=(3,1), flag=wx.ALIGN_LEFT, border=5)
        vbox_grid.Add(self.radio_zh, pos=(3,2), flag=wx.ALIGN_LEFT, border=5)

        vbox_grid.Add(label_about, pos=(4,0), flag=wx.ALIGN_RIGHT, border=5)
        vbox_grid.Add(self.text_multi_text, pos=(4,1), span=(4,3), flag=wx.ALIGN_LEFT, border=5)

        vbox.Add((20,20))
        vbox.Add(vbox_grid, 0, wx.ALIGN_CENTER)
        vbox.Add((20,20))
        vbox.Add(button_close, 0, flag=wx.ALIGN_CENTER)

        panel.SetSizer(vbox)
        panel.Layout()

        self.Bind(wx.EVT_BUTTON, self.OnClose, button_close)
        self.Bind(wx.EVT_CLOSE, self.OnClose)

        self.updateConfigUI()
        self.text_multi_text.LoadFile(dbfilenames.about_file)

    def OnClose(self, event):
        # self.Close()
        dbfilenames.ConfigSave()
        self.Destroy()

    def OnRadioLanguage(self, event):
        radioSelected = event.GetEventObject()
        text = radioSelected.GetLabel()
        dbfilenames.current["lang"] = text

    def OnRadioConvertTo(self, event):
        radioSelected = event.GetEventObject()
        text = radioSelected.GetLabel()
        dbfilenames.current["convert_to"] = text

    def OnRadioFormat(self, event):
        radioSelected = event.GetEventObject()
        text = radioSelected.GetLabel()
        dbfilenames.current["image_format"] = text

    def OnRadioQuality(self, event):
        radioSelected = event.GetEventObject()
        text = radioSelected.GetLabel()
        dbfilenames.current["image_quality"] = text

    def updateConfigUI(self):
        """update current dict to config window"""
        if dbfilenames.current["convert_to"] == "Picture":
            self.radio_convert_picture.SetValue(True)
        elif dbfilenames.current["convert_to"] == "PDF":
            self.radio_convert_pdf.SetValue(True)

        if dbfilenames.current["lang"] == "English":
            self.radio_en.SetValue(True)
        elif dbfilenames.current["lang"] == "中文":
            self.radio_zh.SetValue(True)

        if dbfilenames.current["image_quality"] == "Good":
            self.radio_good.SetValue(True)
        elif dbfilenames.current["image_quality"] == "Better":
            self.radio_better.SetValue(True)
        elif dbfilenames.current["image_quality"] == "Best":
            self.radio_best.SetValue(True)

        if dbfilenames.current["image_format"] == "jpg":
            self.radio_jpg.SetValue(True)
        elif dbfilenames.current["image_format"] == "png":
            self.radio_png.SetValue(True)

dbfilenames = DBfilenames()

class App(wx.App):
    """Application class."""
    def OnInit(self):
        dbfilenames.ConfigLoad()
        # localedir = os.path.join(dbmenus.basepath, "locale")
        # if dbmenus.current["lang"] == "English":
        #     langid = wx.LANGUAGE_ENGLISH_US
        # else:
        #     langid = wx.LANGUAGE_CHINESE_SIMPLIFIED
        # domain = "newmenus"             # the translation file is messages.mo
        # L.Init(langid)
        # L.AddCatalogLookupPathPrefix(localedir)
        # L.AddCatalog(domain)

        self.frame = MainFrame()
        self.frame.Show(True)
        self.SetTopWindow(self.frame)
        return True

def main():
    app = App()
    app.MainLoop()

if __name__ == '__main__':
    main()
