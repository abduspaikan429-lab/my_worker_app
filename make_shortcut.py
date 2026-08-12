import os
import subprocess

desktop = os.path.join(os.path.expanduser("~"), "Desktop")
shortcut_path = os.path.join(desktop, "建筑劳务综合管理平台.lnk")
target_bat = os.path.abspath("启动系统.bat")
work_dir = os.path.abspath(".")

vbs_content = f'''Set WshShell = CreateObject("WScript.Shell")
Set oLink = WshShell.CreateShortcut("{shortcut_path}")
oLink.TargetPath = "{target_bat}"
oLink.WorkingDirectory = "{work_dir}"
oLink.Description = "建筑劳务综合管理平台一键启动"
oLink.Save
'''

temp_vbs = os.path.abspath("temp_shortcut.vbs")
with open(temp_vbs, "w", encoding="gbk") as f:
    f.write(vbs_content)

subprocess.run(f'cscript //Nologo "{temp_vbs}"', shell=True)

if os.path.exists(temp_vbs):
    os.remove(temp_vbs)

print("桌面快捷方式已成功更新！")
