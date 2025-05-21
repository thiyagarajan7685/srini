import wx
import wx.media
import os
from moviepy import VideoFileClip, TextClip, CompositeVideoClip
from threading import Thread
from moviepy.video.io.ffmpeg_tools import ffmpeg_extract_subclip
from moviepy.video.io.ffmpeg_writer import FFMPEG_VideoWriter
from moviepy.tools import extensions_dict
from proglog import ProgressBarLogger


class ExportProgressDialog(wx.ProgressDialog):
    def __init__(self, parent, title, message):
        super().__init__(title, message, maximum=100,
                         parent=parent,
                         style=wx.PD_APP_MODAL | wx.PD_ELAPSED_TIME | wx.PD_REMAINING_TIME)

    def update_progress(self, percent):
        self.Update(percent)

class WxLogger(ProgressBarLogger):
    def __init__(self, progress_dialog):
        super().__init__()
        self.progress_dialog = progress_dialog
        self.last_percent = -1

    def bars_callback(self, bar, attr, value, old_value=None):
        if bar == 'frame_index' and attr == 'index':
            total = self.state['bars'][bar]['total']
            if total:
                percent = int((value / total) * 100)
                if percent != self.last_percent:
                    self.last_percent = percent
                    wx.CallAfter(self.progress_dialog.update_progress, percent)


class VideoFileDropTarget(wx.FileDropTarget):
    def __init__(self, callback):
        super().__init__()
        self.callback = callback

    def OnDropFiles(self, x, y, filenames):
        for filepath in filenames:
            if filepath.lower().endswith(('.mp4', '.avi', '.mov', '.mkv', '.flv', '.wmv')):
                self.callback(filepath)
                return True
        wx.MessageBox("Unsupported file format.", "Error", wx.ICON_ERROR)
        return False


class VideoEditor(wx.Frame):
    def __init__(self, parent, title):
        super(VideoEditor, self).__init__(parent, title=title, size=(820, 650))

        self.panel = wx.Panel(self)
        self.vbox = wx.BoxSizer(wx.VERTICAL)

        self.media_ctrl = wx.media.MediaCtrl(self.panel, style=wx.SIMPLE_BORDER)
        self.vbox.Add(self.media_ctrl, flag=wx.EXPAND | wx.ALL, border=5, proportion=1)

        hbox_buttons = wx.BoxSizer(wx.HORIZONTAL)
        self.load_button = wx.Button(self.panel, label="Load Video")
        self.play_pause_button = wx.Button(self.panel, label="Play")

        self.preview_button = wx.Button(self.panel, label="Preview")
        hbox_buttons.AddMany([(self.load_button, 0, wx.RIGHT, 5),
                              (self.preview_button, 0)])
        self.vbox.Add(hbox_buttons, flag=wx.ALIGN_CENTER | wx.BOTTOM, border=10)

        editor_box = wx.StaticBox(self.panel, label="Editor Options")
        editor_sizer = wx.StaticBoxSizer(editor_box, wx.VERTICAL)

        self.only_audio_cb = wx.CheckBox(self.panel, label="Only Audio")
        self.remove_audio_cb = wx.CheckBox(self.panel, label="Remove Audio")

        text_label = wx.StaticText(self.panel, label="Text:")
        self.text_input = wx.TextCtrl(self.panel)

        text_pos_label = wx.StaticText(self.panel, label="Text Position:")
        self.text_pos_choice = wx.ComboBox(self.panel, choices=["Center", "Top", "Bottom", "Left", "Right"], style=wx.CB_READONLY)
        self.text_pos_choice.SetValue("Center")

        text_color_label = wx.StaticText(self.panel, label="Text Color:")
        self.text_color_choice = wx.ComboBox(self.panel, choices=["White", "Black", "Red", "Blue", "Green"], style=wx.CB_READONLY)
        self.text_color_choice.SetValue("White")

        font_label = wx.StaticText(self.panel, label="Font:")
        self.font_choice = wx.ComboBox(self.panel, choices=["Arial", "Verdana", "Times New Roman", "Courier New"], style=wx.CB_READONLY)
        self.font_choice.SetValue("Arial")

        format_label = wx.StaticText(self.panel, label="Export Format:")
        self.format_choice = wx.ComboBox(self.panel, choices=["mp4", "avi", "mov"], style=wx.CB_READONLY)
        self.format_choice.SetValue("mp4")

        start_time_label = wx.StaticText(self.panel, label="Start Time (s):")
        self.start_time_input = wx.TextCtrl(self.panel, value="0")

        end_time_label = wx.StaticText(self.panel, label="End Time (s):")
        self.end_time_input = wx.TextCtrl(self.panel, value="0")

        # Inside __init__ after self.vbox.Add(hbox_buttons,...)
        self.slider = wx.Slider(self.panel, value=0, minValue=0, maxValue=100, style=wx.SL_HORIZONTAL)

        # Slider with time labels
        slider_hbox = wx.BoxSizer(wx.HORIZONTAL)
        self.current_time_label = wx.StaticText(self.panel, label="00:00")
        self.total_time_label = wx.StaticText(self.panel, label="00:00")

        slider_hbox.Add(self.current_time_label, flag=wx.ALIGN_CENTER_VERTICAL | wx.RIGHT, border=5)
        slider_hbox.Add(self.slider, proportion=1, flag=wx.EXPAND)
        slider_hbox.Add(self.total_time_label, flag=wx.ALIGN_CENTER_VERTICAL | wx.LEFT, border=5)

        self.vbox.Add(slider_hbox, flag=wx.EXPAND | wx.LEFT | wx.RIGHT | wx.BOTTOM, border=10)


        # Output + Play/Pause section
        footer_hbox = wx.BoxSizer(wx.HORIZONTAL)
        self.output_label = wx.StaticText(self.panel, label="Output_video")
        footer_hbox.Add(self.output_label, flag=wx.ALIGN_CENTER_VERTICAL | wx.LEFT, border=10)

        footer_hbox.AddStretchSpacer()
        footer_hbox.Add(self.play_pause_button, flag=wx.RIGHT, border=10)


        self.vbox.Add(footer_hbox, flag=wx.EXPAND | wx.LEFT | wx.RIGHT | wx.BOTTOM, border=10)



        self.export_button = wx.Button(self.panel, label="Export Results")

        grid = wx.FlexGridSizer(0, 2, 5, 10)
        grid.AddMany([
            (self.only_audio_cb), (self.remove_audio_cb),
            (text_label), (self.text_input),
            (text_pos_label), (self.text_pos_choice),
            (text_color_label), (self.text_color_choice),
            (font_label), (self.font_choice),
            (format_label), (self.format_choice),
            (start_time_label), (self.start_time_input),
            (end_time_label), (self.end_time_input),
        ])
        editor_sizer.Add(grid, flag=wx.ALL, border=10)
        editor_sizer.Add(self.export_button, flag=wx.ALIGN_CENTER | wx.ALL, border=10)

        self.vbox.Add(editor_sizer, flag=wx.EXPAND | wx.LEFT | wx.RIGHT | wx.BOTTOM, border=10)

        self.panel.SetSizer(self.vbox)

        self.Bind(wx.EVT_BUTTON, self.on_load_video, self.load_button)
        self.Bind(wx.EVT_BUTTON, self.toggle_play_pause, self.play_pause_button)

        self.Bind(wx.EVT_BUTTON, self.on_export, self.export_button)
        self.Bind(wx.EVT_BUTTON, self.on_preview, self.preview_button)
        self.Bind(wx.EVT_CLOSE, self.on_close)
        self.Bind(wx.EVT_CHECKBOX, self.toggle_fields_for_audio, self.only_audio_cb)
        self.slider.Bind(wx.EVT_SCROLL, self.on_seek)
        self.timer = wx.Timer(self)
        self.Bind(wx.EVT_TIMER, self.on_update_slider, self.timer)

        self.panel.Bind(wx.EVT_DROP_FILES, self.on_drop_files)
        self.panel.SetDropTarget(VideoFileDropTarget(self.load_video_from_path))

        self.video_path = None
        self.video_clip = None
        self.modified_clip = None

        self.Show()

    def load_video_from_path(self, path):
        self.video_path = path
        if self.media_ctrl.Load(path):
            self.media_ctrl.GetBestVirtualSize()
            self.media_ctrl.SetSize(self.media_ctrl.GetBestSize())
            self.panel.Layout()
            self.video_clip = VideoFileClip(self.video_path)
        else:
            wx.MessageBox("Unable to load video.", "Error", wx.ICON_ERROR)


    def toggle_fields_for_audio(self, event):
        only_audio = self.only_audio_cb.GetValue()
        for label in ["Text:", "Text Position:", "Text Color:", "Font:","Export Format:"]:
            widget = self.FindWindowByLabel(label)
            if widget:
                widget.Show(not only_audio)

        self.text_input.Show(not only_audio)
        self.text_pos_choice.Show(not only_audio)
        self.text_color_choice.Show(not only_audio)
        self.font_choice.Show(not only_audio)
        self.remove_audio_cb.Show(not only_audio)
        self.format_choice.Show(not only_audio)
        self.panel.Layout()

    def on_load_video(self, event):
        with wx.FileDialog(self, "Open Video File", wildcard=(
            "Video files (*.mp4;*.avi;*.mov;*.mkv;*.flv;*.wmv)|*.mp4;*.avi;*.mov;*.mkv;*.flv;*.wmv"
        ), style=wx.FD_OPEN | wx.FD_FILE_MUST_EXIST) as file_dialog:
            if file_dialog.ShowModal() == wx.ID_OK:
                self.load_video(file_dialog.GetPath())

    def on_drop_files(self, event):
        paths = event.GetFiles()
        if paths:
            self.load_video(paths[0])

    def load_video(self, path):
        self.video_path = path
        if self.media_ctrl.Load(path):
            self.media_ctrl.GetBestVirtualSize()
            self.media_ctrl.SetSize(self.media_ctrl.GetBestSize())
            self.panel.Layout()
            self.video_clip = VideoFileClip(path)
        else:
            wx.MessageBox("Unable to load video.", "Error", wx.ICON_ERROR)

    def toggle_play_pause(self, event):
        if self.media_ctrl.GetState() == wx.media.MEDIASTATE_PLAYING:
            self.media_ctrl.Pause()
            self.timer.Stop()
            self.play_pause_button.SetLabel("Play")
        else:
            self.media_ctrl.Play()
            self.timer.Start(500)
            self.play_pause_button.SetLabel("Pause")


    def on_preview(self, event):
        if self.modified_clip:
            self.modified_clip.preview()

    def on_export(self, event):
        if not self.video_clip:
            wx.MessageBox("Please load a video first!", "Error", wx.ICON_ERROR)
            return

        progress_dialog = ExportProgressDialog(self, "Exporting", "Please wait while exporting...")

        def export_thread():
            text = self.text_input.GetValue()
            color = self.text_color_choice.GetValue().lower()
            position = self.text_pos_choice.GetValue().lower()
            font = self.font_choice.GetValue()
            fmt = self.format_choice.GetValue()
            start_time = int(self.start_time_input.GetValue())
            end_time = int(self.end_time_input.GetValue())

            clip = self.video_clip.subclipped(start_time, end_time)
            clips = [clip]

            if text:
                txt_clip = TextClip(text=text, font=f"C:/Windows/Fonts/{font.lower()}.ttf", font_size=70, color=color)
                txt_clip = txt_clip.with_position(position).with_start(0).with_duration(end_time - start_time)
                clips.append(txt_clip)

            self.modified_clip = CompositeVideoClip(clips)
            logger = WxLogger(progress_dialog)

            if self.remove_audio_cb.GetValue():
                self.modified_clip = self.modified_clip.without_audio()
            elif self.only_audio_cb.GetValue():
                try:
                    self.modified_clip.audio.write_audiofile("only_audio.mp3",logger=logger)
                    wx.CallAfter(wx.MessageBox, "Exported audio file only.", "Info", wx.ICON_INFORMATION)
                except Exception as e:
                    wx.CallAfter(wx.MessageBox, f"Export failed: {str(e)}", "Error", wx.ICON_ERROR)
                finally:
                    wx.CallAfter(progress_dialog.Destroy)
                return

            output_path = os.path.join(os.getcwd(), f"exported_video.{fmt}")


            try:
                self.modified_clip.write_videofile(output_path, codec="libx264", logger=logger,threads=4)

                wx.CallAfter(wx.MessageBox, f"Exported to {output_path}", "Success", wx.ICON_INFORMATION)
            except Exception as e:
                wx.CallAfter(wx.MessageBox, f"Export failed: {str(e)}", "Error", wx.ICON_ERROR)
            finally:
                wx.CallAfter(progress_dialog.Destroy)

        Thread(target=export_thread).start()

    def on_close(self, event):
        if self.modified_clip:
            self.modified_clip.close()
        self.Destroy()

    def on_update_slider(self, event):
        if self.media_ctrl.Length() > 0:
            pos = self.media_ctrl.Tell()
            duration = self.media_ctrl.Length()
            self.slider.SetRange(0, duration)
            self.slider.SetValue(pos)

            current_time = self.format_time(pos // 1000)
            total_time = self.format_time(duration // 1000)
            self.current_time_label.SetLabel(current_time)
            self.total_time_label.SetLabel(total_time)


    def on_seek(self, event):
        if self.media_ctrl.Length() > 0:
            value = self.slider.GetValue()
            self.media_ctrl.Seek(value)

    def format_time(self, seconds):
        mins = seconds // 60
        secs = seconds % 60
        return f"{mins:02}:{secs:02}"

def main():
    app = wx.App(False)
    VideoEditor(None, title="Video Editor")
    app.MainLoop()


if __name__ == "__main__":
    main()
