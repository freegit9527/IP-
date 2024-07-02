import wx
import wx.lib.filebrowsebutton as FBB
import wx.lib.newevent as newevent
from data_processing import match_data
import pandas as pd

# 定义自定义事件
FilePickedEventType, EVT_FILE_PICKED = newevent.NewEvent()

class FilePickedEvent(wx.PyEvent):
    def __init__(self, path):
        wx.PyEvent.__init__(self)
        self.path = path

class MyFileBrowseButton(FBB.FileBrowseButton):
    def __init__(self, *args, **kwargs):
        super(MyFileBrowseButton, self).__init__(*args, **kwargs)
        self.Bind(wx.EVT_BUTTON, self.onButton)

    def onButton(self, event):
        path = self.GetValue()
        wx.PostEvent(self.GetParent(), FilePickedEvent(path=path))

class IPMergerGUI(wx.Frame):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, title="IP & String Merger Tool", size=(600, 400))
        self.InitUI()

    def InitUI(self):
        panel = wx.Panel(self)
        vbox = wx.BoxSizer(wx.VERTICAL)
        
        # 文件选择和列选择部分
        self.file1_picker = MyFileBrowseButton(panel, labelText="File 1:")
        self.col1_spinner = wx.SpinCtrl(panel, min=1, max=100)
        self.col1_type_radio_ip = wx.RadioButton(panel, label='IP', style=wx.RB_GROUP)
        self.col1_type_radio_str = wx.RadioButton(panel, label='String')
        
        self.file2_picker = MyFileBrowseButton(panel, labelText="File 2:")
        self.col2_spinner = wx.SpinCtrl(panel, min=1, max=100)
        self.col2_type_radio_ip = wx.RadioButton(panel, label='IP', style=wx.RB_GROUP)
        self.col2_type_radio_str = wx.RadioButton(panel, label='String')
        
        vbox.Add(self.create_file_col_row(panel, self.file1_picker, self.col1_spinner, self.col1_type_radio_ip, self.col1_type_radio_str), proportion=0, flag=wx.EXPAND|wx.LEFT|wx.RIGHT|wx.TOP, border=10)
        vbox.Add(self.create_file_col_row(panel, self.file2_picker, self.col2_spinner, self.col2_type_radio_ip, self.col2_type_radio_str), proportion=0, flag=wx.EXPAND|wx.LEFT|wx.RIGHT|wx.TOP, border=10)
        
        # 合并按钮
        merge_button = wx.Button(panel, label="Merge and Save")
        merge_button.Bind(wx.EVT_BUTTON, self.on_merge)
        vbox.Add(merge_button, proportion=0, flag=wx.ALIGN_CENTER|wx.TOP|wx.BOTTOM, border=10)
        
        # 使用说明
        instructions = """
        Usage Instructions:
        1. Select your input files using the "Browse" buttons.
        2. Choose the column numbers and type ('IP' or 'String') for comparison.
        3. Click "Merge and Save" to process and save the output.
        """
        instructions_text = wx.StaticText(panel, label=instructions, style=wx.ALIGN_LEFT)
        instructions_text.Wrap(550)  # 自动换行
        vbox.Add(instructions_text, proportion=0, flag=wx.LEFT|wx.RIGHT|wx.BOTTOM|wx.EXPAND, border=10)
        
        panel.SetSizer(vbox)
        self.Centre()
        self.Show()

    def create_file_col_row(self, panel, file_picker, col_spinner, radio_ip, radio_str):
        hbox = wx.BoxSizer(wx.HORIZONTAL)
        hbox.Add(file_picker, proportion=1, flag=wx.EXPAND|wx.ALL, border=5)
        hbox.Add(wx.StaticText(panel, label="Column:"), flag=wx.ALIGN_CENTER_VERTICAL|wx.LEFT, border=5)
        hbox.Add(col_spinner, proportion=0, flag=wx.EXPAND|wx.LEFT|wx.RIGHT, border=5)
        hbox.Add(radio_ip, flag=wx.LEFT, border=10)
        hbox.Add(radio_str, flag=wx.LEFT, border=10)
        return hbox

    def on_merge(self, event):
        file1_path = self.file1_picker.GetValue()
        col1 = self.col1_spinner.GetValue()
        col1_type = 'ip' if self.col1_type_radio_ip.GetValue() else 'string'
        
        file2_path = self.file2_picker.GetValue()
        col2 = self.col2_spinner.GetValue()
        col2_type = 'ip' if self.col2_type_radio_ip.GetValue() else 'string'
        
        try:
            result_df = match_data(file1_path, col1_type, col1-1, file2_path, col2_type, col2-1)
            
            # 使用wx.FileDialog让用户选择输出文件的位置
            with wx.FileDialog(self, "Save merged file as...", wildcard="Excel files (*.xlsx)|*.xlsx", style=wx.FD_SAVE | wx.FD_OVERWRITE_PROMPT) as fileDialog:
                if fileDialog.ShowModal() == wx.ID_OK:
                    output_path = fileDialog.GetPath()
                    result_df.to_excel(output_path, index=False)
                    wx.MessageBox("Merge completed successfully!", "Success", wx.OK | wx.ICON_INFORMATION)
        except Exception as e:
            wx.LogError(f"Error during merging: {str(e)}")

if __name__ == '__main__':
    app = wx.App(False)
    frame = IPMergerGUI(None)
    app.MainLoop()