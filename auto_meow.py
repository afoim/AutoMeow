import keyboard
import time
import pystray
from PIL import Image, ImageDraw, ImageFont
import threading
import sys

class AutoMeow:
    def __init__(self):
        self.enabled = False
        self.last_time = 0
        self.icon = None
        self.setup_tray()

    def create_icon(self, color):
        image = Image.new('RGB', (32, 32), color)
        draw = ImageDraw.Draw(image)
        try:
            font = ImageFont.truetype("simhei.ttf", 24)
        except:
            font = ImageFont.load_default()
        
        text = "喵"
        left, top, right, bottom = draw.textbbox((0, 0), text, font=font)
        x = (32 - (right - left)) // 2
        y = (32 - (bottom - top)) // 2
        draw.text((x, y), text, font=font, fill='white')
        return image
        
    def toggle(self, *args):
        self.enabled = not self.enabled
        if self.icon:
            color = 'green' if self.enabled else 'red'
            self.icon.icon = self.create_icon(color)
            self.icon.title = f"AutoMeow ({'启用' if self.enabled else '禁用'})"
            
            if not self.enabled:
                keyboard.unhook_all()
            else:
                keyboard.on_press_key("enter", self.on_enter, suppress=True)
    
    def setup_tray(self):
        menu = pystray.Menu(
            pystray.MenuItem("启用/禁用", self.toggle),
            pystray.MenuItem("退出", self.quit_app)
        )
        self.icon = pystray.Icon("AutoMeow", 
                                self.create_icon('red'),
                                "AutoMeow (禁用)", 
                                menu)

    def quit_app(self, *args):
        if self.icon:
            self.icon.stop()
        keyboard.unhook_all()
    
    def on_enter(self, e):
        current_time = time.time()
        if not self.enabled or current_time - self.last_time < 0.1:
            return
        
        self.last_time = current_time
        e.event_type = 'down'
        keyboard.block_key(e.scan_code)
        
        keyboard.write("喵")
        time.sleep(0.05)
        keyboard.press_and_release("enter")
        
        keyboard.unblock_key(e.scan_code)
        return False

    def run(self):
        self.icon.run()

if __name__ == "__main__":
    if getattr(sys, 'frozen', False):
        # 如果是打包后的exe运行
        import os
        os.environ['PYTHONUNBUFFERED'] = '1'
    meow = AutoMeow()
    meow.run()
