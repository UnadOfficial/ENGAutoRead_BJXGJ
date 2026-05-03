import sys
import io
import os

os.environ['PYTHONUNBUFFERED'] = '1'

if sys.platform == 'win32':
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', line_buffering=True)
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', line_buffering=True)

import subprocess
import time
import re
import json
from PIL import Image
import pytesseract
from tts_engine import TTSEngineFactory


class EnhancedWordAutomator:
    def __init__(self, adb_path="adb", config_file="config.json"):
        self.adb_path = adb_path
        self.device_id = None
        self.script_dir = os.path.dirname(os.path.abspath(__file__))
        self.config = self.load_config(config_file)
        self.tts_engines = TTSEngineFactory.create_fallback_chain(self.config)
        self._initialize_tts_engines()
        self.stats = {'total': 0, 'success': 0, 'failed': 0, 'ocr_failed': 0, 'click_failed': 0}

    def load_config(self, config_file):
        default_config = {
            "microphone_position": {"x": 540, "y": 1580},
            "stop_button_position": {"x": 540, "y": 1580},
            "continue_button_position": {"x": 620, "y": 2640},
            "word_region": {"top": 760, "bottom": 1495, "left": 38, "right": 1225},
            "tesseract": {"path": ""},
            "tts_settings": {
                "engine": "kokoro",
                "speed": 1.0,
                "kokoro_voice": "zm_yunjian"
            },
            "recording_settings": {"duration": 5.0, "post_recording_delay": 3.0},
            "debug_settings": {"save_screenshots": True, "verbose_config": False},
            "screenshot_delay": 0.5,
            "tts_duration": 2.5,
            "click_delay": 0.3
        }

        if os.path.exists(config_file):
            try:
                with open(config_file, 'r', encoding='utf-8') as f:
                    user_config = json.load(f)
                    default_config.update(user_config)
                    print(f"[CONFIG] 已加载配置文件: {config_file}")
                    print(f"[CONFIG] TTS引擎: {default_config.get('tts_settings', {}).get('engine', 'N/A')}")
                    print(f"[CONFIG] Kokoro音色: {default_config.get('tts_settings', {}).get('kokoro_voice', 'N/A')}")
                    if default_config.get('debug_settings', {}).get('verbose_config', False):
                        print(f"[CONFIG] 完整配置: {json.dumps(default_config, indent=2, ensure_ascii=False)}")
            except Exception as e:
                print(f"[CONFIG] 加载配置失败: {e}")
        else:
            with open(config_file, 'w', encoding='utf-8') as f:
                json.dump(default_config, f, indent=4, ensure_ascii=False)
            print(f"[CONFIG] 已创建默认配置文件: {config_file}")

        tesseract_path = default_config.get('tesseract', {}).get('path', '')
        if tesseract_path:
            pytesseract.pytesseract.tesseract_cmd = tesseract_path

        return default_config

    def _initialize_tts_engines(self):
        for engine in self.tts_engines:
            engine.initialize()

    def speak(self, word):
        for i, engine in enumerate(self.tts_engines):
            try:
                if engine.is_available:
                    print(f"[TTS] 使用引擎: {engine.__class__.__name__}")
                    engine.speak(word)
                    return
            except Exception as e:
                print(f"[WARN] {engine.__class__.__name__} 失败: {e}")
                if i < len(self.tts_engines) - 1:
                    print("[TTS] 尝试回退到下一个引擎...")
                else:
                    print("[ERR] 所有 TTS 引擎都失败，跳过朗读")

    def execute_adb(self, command, timeout=15, text_mode=True):
        full_command = f"{self.adb_path} {command}"
        try:
            if text_mode:
                result = subprocess.run(full_command, shell=True, capture_output=True,
                                       text=True, encoding='utf-8', errors='ignore', timeout=timeout)
                return result.stdout.strip(), result.stderr.strip(), result.returncode
            else:
                result = subprocess.run(full_command, shell=True, capture_output=True, timeout=timeout)
                return result.stdout, result.stderr, result.returncode
        except subprocess.TimeoutExpired:
            return ("", "Timeout", -1) if text_mode else (b"", b"Timeout", -1)
        except Exception as e:
            return ("", str(e), -1) if text_mode else (b"", str(e).encode('utf-8'), -1)

    def connect_device(self):
        print("正在连接设备...")
        stdout, stderr, code = self.execute_adb("devices")
        if code != 0:
            print(f"连接失败: {stderr}")
            return False
        lines = stdout.split('\n')
        devices = [line.split()[0] for line in lines[1:] if line.strip() and 'device' in line]
        if not devices:
            print("未找到设备，请检查模拟器是否启动、USB调试是否开启")
            return False
        self.device_id = devices[0]
        print(f"[OK] 已连接设备: {self.device_id}")
        return True

    def capture_screen(self):
        time.sleep(self.config.get('screenshot_delay', 0.5))
        stdout, stderr, code = self.execute_adb("exec-out screencap -p", text_mode=False)
        if code != 0 or not stdout:
            if isinstance(stderr, bytes):
                stderr = stderr.decode('utf-8', errors='ignore')
            raise Exception(f"截图失败: {stderr}")
        return Image.open(io.BytesIO(stdout))

    def crop_word_region(self, image):
        r = self.config['word_region']
        return image.crop((r['left'], r['top'], r['right'], r['bottom']))

    def extract_word_from_image(self, image):
        try:
            text = pytesseract.image_to_string(image, config=r'--oem 3 --psm 6 -l eng+chi_sim').strip()
            print(f"[OCR] 原始识别: {text}")
            text = re.sub(r'[\u4e00-\u9fff]', '', text)
            text = re.sub(r'[^a-zA-Z0-9\s.,!?;:\'"()\[\]{}+\-*/=<>^_~]', '', text)
            text = re.sub(r'\s+', ' ', text).strip()
            print(f"[OCR] 过滤后: {text}")
            return text if text else None
        except Exception as e:
            print(f"[WARN] OCR识别失败: {e}")
            return None

    def tap_screen(self, x, y):
        stdout, stderr, code = self.execute_adb(f"shell input tap {int(x)} {int(y)}")
        time.sleep(self.config.get('click_delay', 0.3))
        return code == 0

    def click_microphone(self):
        pos = self.config['microphone_position']
        print(f"[MIC] 点击麦克风 ({pos['x']}, {pos['y']})")
        success = self.tap_screen(pos['x'], pos['y'])
        print("[OK] 麦克风点击成功" if success else "[ERR] 麦克风点击失败")
        return success

    def click_stop_button(self):
        pos = self.config.get('stop_button_position', self.config['microphone_position'])
        print(f"[STOP] 点击停止 ({pos['x']}, {pos['y']})")
        success = self.tap_screen(pos['x'], pos['y'])
        print("[OK] 停止点击成功" if success else "[ERR] 停止点击失败")
        return success

    def click_continue_button(self):
        pos = self.config.get('continue_button_position', self.config['microphone_position'])
        print(f"[CONTINUE] 点击继续 ({pos['x']}, {pos['y']})")
        success = self.tap_screen(pos['x'], pos['y'])
        print("[OK] 继续点击成功" if success else "[ERR] 继续点击失败")
        return success

    def process_single_word(self):
        self.stats['total'] += 1
        try:
            print(f"\n{'='*60}")
            print(f"[LOG] 开始处理第 {self.stats['total']} 个单词")
            print(f"{'='*60}")

            save_screenshots = self.config.get('debug_settings', {}).get('save_screenshots', True)

            # 1. 截图
            print("[LOG] [1/9] 截取屏幕...")
            image = self.capture_screen()
            if save_screenshots:
                image.save(os.path.join(self.script_dir, "debug_full_screenshot.png"))

            # 2. 裁剪
            print("[LOG] [2/9] 裁剪单词区域...")
            word_image = self.crop_word_region(image)
            if save_screenshots:
                word_image.save(os.path.join(self.script_dir, "debug_word_region.png"))

            # 3. OCR
            print("[LOG] [3/9] OCR识别...")
            word = self.extract_word_from_image(word_image)
            if not word:
                print("[LOG] [WARN] 未识别到有效单词，跳过")
                self.stats['ocr_failed'] += 1
                self.stats['failed'] += 1
                return False
            print(f"[LOG] [3/9] 识别完成: {word}")

            # 4. 点击麦克风
            print("[LOG] [4/9] 点击麦克风...")
            if not self.click_microphone():
                self.stats['click_failed'] += 1
                self.stats['failed'] += 1
                return False

            # 5. TTS朗读
            print("[LOG] [5/9] TTS朗读...")
            self.speak(word)

            # 6. 等待录音
            duration = self.config.get('recording_settings', {}).get('duration', 5.0)
            print(f"[LOG] [6/9] 等待录音 {duration} 秒...")
            time.sleep(duration)

            # 7. 点击停止
            print("[LOG] [7/9] 点击停止...")
            self.click_stop_button()

            # 8. 等待
            post_delay = self.config.get('recording_settings', {}).get('post_recording_delay', 3.0)
            print(f"[LOG] [8/9] 等待 {post_delay} 秒...")
            time.sleep(post_delay)

            # 9. 点击继续
            print("[LOG] [9/9] 点击继续...")
            self.click_continue_button()

            self.stats['success'] += 1
            print(f"\n[LOG] ✓ 处理完成！成功: {self.stats['success']}/{self.stats['total']}")
            return True

        except Exception as e:
            print(f"[ERR] 处理出错: {e}")
            self.stats['failed'] += 1
            return False

    def print_statistics(self):
        print(f"\n{'='*60}")
        print("   处理统计")
        print(f"{'='*60}")
        print(f"  总处理: {self.stats['total']}")
        print(f"  成功:   {self.stats['success']}")
        print(f"  失败:   {self.stats['failed']}")
        if self.stats['total'] > 0:
            print(f"  成功率: {self.stats['success']/self.stats['total']*100:.1f}%")
        print(f"  OCR失败: {self.stats['ocr_failed']}")
        print(f"  点击失败: {self.stats['click_failed']}")
        print(f"{'='*60}\n")

    def run_automation(self, mode="single", interval=5, max_count=None):
        print(f"\n{'='*60}")
        print("   启动单词自动化工具")
        print(f"{'='*60}\n")

        if mode == "single":
            return self.process_single_word()

        count = 0
        print(f"连续模式启动 (间隔: {interval}s)")
        if max_count:
            print(f"   最大次数: {max_count}")
        print("   按 Ctrl+C 停止\n")

        try:
            while True:
                if max_count and count >= max_count:
                    break
                count += 1
                self.process_single_word()
                print(f"等待 {interval} 秒...")
                time.sleep(interval)
        except KeyboardInterrupt:
            print("\n\n用户中断执行")

        print(f"\n总共执行了 {count} 次")
        self.print_statistics()


def main():
    script_dir = os.path.dirname(os.path.abspath(__file__))
    config_path = os.path.join(script_dir, "config.json")

    automator = EnhancedWordAutomator(config_file=config_path)

    if not automator.connect_device():
        return

    print("\n请选择模式：")
    print("  [1] 单次执行")
    print("  [2] 连续循环")
    print("  [3] 校准坐标")

    choice = input("\n输入选项: ").strip()

    if choice == "1":
        automator.run_automation(mode="single")
    elif choice == "2":
        interval = input("循环间隔(秒) [默认5]: ").strip()
        interval = float(interval) if interval else 5
        max_count = input("最大次数 [留空=无限]: ").strip()
        max_count = int(max_count) if max_count else None
        automator.run_automation(mode="continuous", interval=interval, max_count=max_count)
    elif choice == "3":
        print("\n校准模式：连续点击麦克风位置，观察是否准确\n")
        try:
            for i in range(5):
                print(f"第 {i+1} 次点击...")
                automator.click_microphone()
                time.sleep(2)
        except KeyboardInterrupt:
            print("\n停止校准")
    else:
        print("无效选项")


if __name__ == "__main__":
    main()
