import sys
import codecs
import io
import os

# 禁用输出缓冲，确保实时显示日志
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
pytesseract.pytesseract.tesseract_cmd = r'F:\Soft\tocr' 
try:
    import pyttsx3
    HAS_PYTTSX3 = True
except ImportError:
    HAS_PYTTSX3 = False
    print("注意: pyttsx3 未安装，Windows TTS功能将不可用")

# Kokoro-82M TTS 导入
try:
    from kokoro import KPipeline
    import soundfile as sf
    import sounddevice as sd
    import numpy as np
    HAS_KOKORO = True
except ImportError:
    HAS_KOKORO = False
    print("注意: Kokoro-82M 未安装，将使用Windows TTS")


class EnhancedWordAutomator:
    def __init__(self, adb_path="adb", config_file="config.json"):
        self.adb_path = adb_path
        self.device_id = None
        # 保存脚本目录，用于保存调试图片
        self.script_dir = os.path.dirname(os.path.abspath(__file__))
        self.config = self.load_config(config_file)
        self.windows_tts_engine = None
        self.kokoro_pipeline = None
        
        # 统计信息
        self.stats = {
            'total': 0,        # 总处理次数
            'success': 0,      # 成功次数
            'failed': 0,       # 失败次数
            'ocr_failed': 0,   # OCR识别失败次数
            'click_failed': 0  # 点击失败次数
        }
        
        self._init_windows_tts()
        self._init_kokoro_tts()
        
    def load_config(self, config_file):
        """加载配置文件"""
        default_config = {
            "microphone_position": {"x": 540, "y": 1580},
            "stop_button_position": {"x": 540, "y": 1580},
            "continue_button_position": {"x": 620, "y": 2640},
            "word_region": {
                "top": 760,    
                "bottom": 1495, 
                "left": 38, "right": 1225
            },
            "tesseract": {
                "path": "F:\\Soft\\tocr\\tesseract.exe"
            },
            "tts_settings": {
                "engine": "kokoro",
                "_comment": "TTS引擎选择：'windows'使用Windows系统TTS，'kokoro'使用Kokoro-82M",
                "language": "en-US",
                "speed": 1.0,
                "pitch": 1.0,
                "kokoro_voice": "zm_yunjian",
                "_comment2": "Kokoro音色：zf_xiaobei,zf_xiaoni,zf_xiaoxiao,zf_xiaoyi,zm_yunjian,zm_yunxi,zm_yunxia,zm_yunyang"
            },
            "recording_settings": {
                "duration": 10.0
            },
            "screenshot_delay": 0.5,
            "tts_duration": 2.5,
            "click_delay": 0.3
        }
        
        print(f"[CONFIG] 配置文件路径: {os.path.abspath(config_file)}")
        print(f"[CONFIG] 配置文件是否存在: {os.path.exists(config_file)}")
        
        if os.path.exists(config_file):
            try:
                with open(config_file, 'r', encoding='utf-8') as f:
                    user_config = json.load(f)
                    default_config.update(user_config)
                    print(f"[CONFIG] 已加载配置文件: {config_file}")
                    
                    # 根据verbose_config决定是否显示详细配置
                    verbose = default_config.get('debug_settings', {}).get('verbose_config', False)
                    if verbose:
                        print(f"[CONFIG] 读取到的用户配置: {json.dumps(user_config, indent=2, ensure_ascii=False)}")
                        print(f"[CONFIG] 最终完整配置: {json.dumps(default_config, indent=2, ensure_ascii=False)}")
                    
                    print(f"[CONFIG] TTS引擎: {default_config.get('tts_settings', {}).get('engine', 'N/A')}")
                    print(f"[CONFIG] Kokoro音色: {default_config.get('tts_settings', {}).get('kokoro_voice', 'N/A')}")
            except Exception as e:
                print(f"[CONFIG] 加载配置失败，使用默认配置: {e}")
                import traceback
                traceback.print_exc()
        else:
            # 创建默认配置文件
            with open(config_file, 'w', encoding='utf-8') as f:
                json.dump(default_config, f, indent=4, ensure_ascii=False)
            print(f"[CONFIG] 已创建默认配置文件: {config_file}")
                
        # 设置Tesseract路径
        tesseract_path = default_config.get('tesseract', {}).get('path', '')
        if tesseract_path:
            pytesseract.pytesseract.tesseract_cmd = tesseract_path
            print(f"[OK] Tesseract路径已设置: {tesseract_path}")
                
        return default_config
    
    def _init_windows_tts(self):
        """初始化Windows TTS引擎"""
        if not HAS_PYTTSX3:
            return
            
        try:
            self.windows_tts_engine = pyttsx3.init()
            
            # 设置TTS参数
            tts_cfg = self.config.get('tts_settings', {})
            rate = tts_cfg.get('speed', 1.0)
            
            # pyttsx3的rate范围大约是0-200，默认200
            # 我们把1.0映射到150左右
            engine_rate = int(150 * rate)
            self.windows_tts_engine.setProperty('rate', engine_rate)
            
            # 尝试设置英文语音
            voices = self.windows_tts_engine.getProperty('voices')
            for voice in voices:
                if 'english' in voice.name.lower() or 'en' in voice.id.lower():
                    self.windows_tts_engine.setProperty('voice', voice.id)
                    break
            
            print("[OK] Windows TTS引擎初始化成功")
            
        except Exception as e:
            print(f"[WARN] Windows TTS初始化失败: {e}")
            self.windows_tts_engine = None
    
    def _init_kokoro_tts(self):
        """初始化Kokoro-82M TTS引擎（延迟初始化）"""
        if not HAS_KOKORO:
            self.kokoro_available = False
            return
            
        self.kokoro_available = False
        self.kokoro_initialized = False
        # 不在这里立即初始化，第一次使用时再初始化
        print("[Kokoro] Kokoro-82M已就绪，首次使用时会自动初始化")
    
    def speak_with_kokoro_tts(self, word):
        """使用Kokoro-82M TTS朗读（带延迟初始化）"""
        if not HAS_KOKORO:
            print("[WARN] Kokoro-82M未安装，使用Windows TTS")
            self.speak_with_windows_tts(word)
            return
            
        # 延迟初始化
        if not self.kokoro_initialized:
            try:
                print("[Kokoro] 正在初始化Kokoro-82M TTS...")
                print("[Kokoro] 首次使用可能需要下载模型（约300MB）...")
                # 使用英文模型
                self.kokoro_pipeline = KPipeline(lang_code='a')
                self.kokoro_initialized = True
                self.kokoro_available = True
                print("[OK] Kokoro-82M TTS初始化成功")
            except Exception as e:
                print(f"[WARN] Kokoro-82M初始化失败: {e}")
                print("[INFO] 可能是内存不足或网络问题")
                print("[INFO] 回退到Windows TTS")
                self.kokoro_available = False
                self.speak_with_windows_tts(word)
                return
        
        if not self.kokoro_available:
            self.speak_with_windows_tts(word)
            return
            
        try:
            tts_cfg = self.config.get('tts_settings', {})
            voice = tts_cfg.get('kokoro_voice', 'zm_yunjian')
            
            print(f"[Kokoro] 正在朗读: {word}")
            print(f"[Kokoro] 使用音色: {voice}")
            
            # 生成语音
            generator = self.kokoro_pipeline(word, voice=voice)
            result = next(generator)
            
            # 获取音频数据
            audio_tensor = result.output.audio
            audio_numpy = audio_tensor.detach().cpu().numpy()
            
            # 处理维度
            if audio_numpy.ndim == 1:
                audio_numpy = audio_numpy.reshape(-1, 1)
            
            # 采样率
            sample_rate = 24000
            
            # 播放音频
            sd.play(audio_numpy, sample_rate)
            sd.wait()
            
            # 等待播放完成
            duration = self.config.get('tts_duration', 2.5)
            time.sleep(duration)
            
        except Exception as e:
            print(f"[WARN] Kokoro-82M朗读失败: {e}")
            print("[INFO] 回退到Windows TTS")
            self.speak_with_windows_tts(word)
    
    def speak_with_windows_tts(self, word):
        """使用Windows系统TTS朗读单词"""
        if not HAS_PYTTSX3 or not self.windows_tts_engine:
            print("[WARN] Windows TTS不可用，跳过朗读")
            return
            
        try:
            print(f"[TTS] 正在朗读: {word}")
            self.windows_tts_engine.say(word)
            self.windows_tts_engine.runAndWait()
            
            # 等待播放完成
            duration = self.config.get('tts_duration', 2.5)
            time.sleep(duration)
            
        except Exception as e:
            print(f"[WARN] Windows TTS朗读失败: {e}")
    
    def execute_adb(self, command, timeout=15, text_mode=True):
        """执行ADB命令"""
        full_command = f"{self.adb_path} {command}"
        try:
            if text_mode:
                result = subprocess.run(
                    full_command,
                    shell=True,
                    capture_output=True,
                    text=True,
                    encoding='utf-8',
                    errors='ignore',
                    timeout=timeout
                )
                return result.stdout.strip(), result.stderr.strip(), result.returncode
            else:
                # 二进制模式（用于截图）
                result = subprocess.run(
                    full_command,
                    shell=True,
                    capture_output=True,
                    timeout=timeout
                )
                return result.stdout, result.stderr, result.returncode
        except subprocess.TimeoutExpired:
            if text_mode:
                return "", "Timeout", -1
            else:
                return b"", b"Timeout", -1
        except Exception as e:
            if text_mode:
                return "", str(e), -1
            else:
                return b"", str(e).encode('utf-8'), -1
    
    def get_screen_size(self):
        """获取屏幕分辨率"""
        stdout, stderr, code = self.execute_adb("shell wm size")
        if code == 0 and stdout:
            match = re.search(r'(\d+)x(\d+)', stdout)
            if match:
                width, height = int(match.group(1)), int(match.group(2))
                return width, height
        return 1080, 1920  # 默认分辨率
    
    def connect_device(self):
        """连接设备"""
        print("正在连接设备...")
        stdout, stderr, code = self.execute_adb("devices")
        
        if code != 0:
            print(f"连接失败: {stderr}")
            return False
            
        lines = stdout.split('\n')
        devices = [line.split()[0] for line in lines[1:] if line.strip() and 'device' in line]
        
        if not devices:
            print("未找到设备，请检查：")
            print("  1. 模拟器是否启动")
            print("  2. USB调试是否开启")
            print("  3. ADB是否正确安装")
            return False
            
        self.device_id = devices[0]
        print(f"[OK] 已连接设备: {self.device_id}")
        return True
    
    def capture_screen(self):
        """截取屏幕并返回PIL Image对象"""
        time.sleep(self.config.get('screenshot_delay', 0.5))
        stdout, stderr, code = self.execute_adb("exec-out screencap -p", text_mode=False)
        
        if code != 0 or not stdout:
            if isinstance(stderr, bytes):
                stderr = stderr.decode('utf-8', errors='ignore')
            raise Exception(f"截图失败: {stderr}")
            
        image = Image.open(io.BytesIO(stdout))
        return image
    
    def crop_word_region(self, image):
        """裁剪出单词显示区域"""
        region = self.config['word_region']
        
        left = region['left']
        top = region['top']
        right = region['right']
        bottom = region['bottom']
        
        cropped = image.crop((left, top, right, bottom))
        return cropped
    
    def extract_word_from_image(self, image):
        """从图像中提取英文（支持单词或句子，排除中文和特殊符号）"""
        try:
            # 使用pytesseract OCR，同时识别中英文
            custom_config = r'--oem 3 --psm 6 -l eng+chi_sim'
            text = pytesseract.image_to_string(image, config=custom_config).strip()
            
            print(f"[OCR] 原始识别: {text}")
            
            # 排除中文字符
            # 中文Unicode范围: \u4e00-\u9fff
            text = re.sub(r'[\u4e00-\u9fff]', '', text)
            
            # 排除特殊符号，只保留：
            # - 英文字母 (a-z, A-Z)
            # - 数字 (0-9)
            # - 常用标点 (.,!?;:'"()[]{}+-*/=<>@#$%^&_~)
            # - 空格
            text = re.sub(r'[^a-zA-Z0-9\s.,!?;:\'"()\[\]{}+\-*/=<>^_~]', '', text)
            
            # 清理多余的空白
            text = re.sub(r'\s+', ' ', text).strip()
            
            print(f"[OCR] 过滤后: {text}")
            
            # 如果能识别到文本，直接返回（支持整句）
            if text and len(text.strip()) > 0:
                return text
            
            return None
            
        except Exception as e:
            print(f"[WARN] OCR识别失败: {e}")
            print("[INFO] 请确保已正确安装Tesseract OCR")
            print("[INFO] 下载地址: https://github.com/UB-Mannheim/tesseract/wiki")
            print("[INFO] 安装后可在config.json中设置tesseract.path")
            return None
    
    def speak_with_system_tts(self, word):
        """根据配置选择TTS引擎朗读"""
        tts_engine = self.config.get('tts_settings', {}).get('engine', 'windows')
        print(f"[TTS] 使用引擎: {tts_engine}")
        
        if tts_engine == 'kokoro':
            self.speak_with_kokoro_tts(word)
        elif tts_engine == 'windows':
            self.speak_with_windows_tts(word)
        elif tts_engine == 'android':
            self.speak_with_android_tts(word)
        else:
            print(f"[WARN] 未知TTS引擎: {tts_engine}，使用Windows TTS")
            self.speak_with_windows_tts(word)
    
    def speak_with_android_tts(self, word):
        """使用Android系统TTS引擎朗读"""
        tts_cfg = self.config['tts_settings']
        language = tts_cfg.get('language', 'en-US')
        
        print(f"[TTS] [Android] 正在朗读: {word}")
        
        # 方法1：通过Intent调用TTS
        commands = [
            # 尝试多种TTS调用方式
            f'am start -a android.speech.tts.engine.CHECK_TTS_DATA',
            f'shell input text "{word}" && am broadcast -a com.android.intent.action.TTS_SPEAK --es msg "{word}"',
        ]
        
        for cmd in commands:
            stdout, stderr, code = self.execute_adb(cmd)
            if code == 0:
                break
        
        # 等待TTS播放完成
        duration = self.config.get('tts_duration', 2.5)
        time.sleep(duration)
    
    def tap_screen(self, x, y):
        """点击屏幕指定坐标"""
        cmd = f"shell input tap {int(x)} {int(y)}"
        stdout, stderr, code = self.execute_adb(cmd)
        time.sleep(self.config.get('click_delay', 0.3))
        return code == 0
    
    def click_microphone(self):
        """点击麦克风按钮（开始录音）"""
        mic_pos = self.config['microphone_position']
        
        x = mic_pos['x']
        y = mic_pos['y']
        
        print(f"[MIC] 点击麦克风按钮 ({int(x)}, {int(y)})")
        success = self.tap_screen(x, y)
        
        if success:
            print("[OK] 麦克风按钮点击成功，开始录音...")
        else:
            print("[ERR] 点击失败")
        
        return success
    
    def click_stop_button(self):
        """点击停止按钮（结束录音）"""
        stop_pos = self.config.get('stop_button_position', 
                                   self.config['microphone_position'])
        
        x = stop_pos['x']
        y = stop_pos['y']
        
        print(f"[STOP] 点击停止按钮 ({int(x)}, {int(y)})")
        success = self.tap_screen(x, y)
        
        if success:
            print("[OK] 停止按钮点击成功，录音结束")
        else:
            print("[ERR] 停止按钮点击失败")
        
        return success
    
    def click_continue_button(self):
        """点击继续按钮（进入下一题）"""
        continue_pos = self.config.get('continue_button_position', 
                                     self.config['microphone_position'])
        
        x = continue_pos['x']
        y = continue_pos['y']
        
        print(f"[CONTINUE] 点击继续按钮 ({int(x)}, {int(y)})")
        success = self.tap_screen(x, y)
        
        if success:
            print("[OK] 继续按钮点击成功，进入下一题")
        else:
            print("[ERR] 继续按钮点击失败")
        
        return success
    
    def wait_for_recording(self):
        """等待录音完成"""
        duration = self.config.get('recording_settings', {}).get('duration', 3.0)
        print(f"[WAIT] 正在录音，等待 {duration} 秒...")
        time.sleep(duration)
    
    def process_single_word(self):
        self.stats['total'] += 1
        
        try:
            print("\n" + "="*60)
            print(f"[LOG] 开始处理第 {self.stats['total']} 个单词")
            print("="*60)
            
            # 检查是否保存调试图片
            save_screenshots = self.config.get('debug_settings', {}).get('save_screenshots', True)
            
            # 1. 截图
            print("[LOG] [1/9] 正在截取屏幕...")
            image = self.capture_screen()
            print("[LOG] [1/9] 截图完成")
            
            # 保存完整截图（调试用）
            if save_screenshots:
                debug_full = os.path.join(self.script_dir, "debug_full_screenshot.png")
                image.save(debug_full)
                print(f"[LOG] [DEBUG] 完整截图已保存: {debug_full}")
            
            # 2. 裁剪单词区域
            print("[LOG] [2/9] 正在裁剪单词区域...")
            word_image = self.crop_word_region(image)
            print("[LOG] [2/9] 裁剪完成")
            
            if save_screenshots:
                debug_word = os.path.join(self.script_dir, "debug_word_region.png")
                word_image.save(debug_word)
                print(f"[LOG] [DEBUG] 单词区域已保存: {debug_word}")
            
            # 3. OCR识别
            print("[LOG] [3/9] 正在OCR识别...")
            word = self.extract_word_from_image(word_image)
            
            if not word:
                print("[LOG] [WARN] 未识别到有效单词，跳过")
                self.stats['ocr_failed'] += 1
                self.stats['failed'] += 1
                print("="*60 + "\n")
                return False
            
            print(f"[LOG] [3/9] OCR识别完成: {word}")
            
            # 4. 点击麦克风按钮（开始录音）
            print("[LOG] [4/9] 正在点击麦克风按钮...")
            mic_success = self.click_microphone()
            if not mic_success:
                print("[LOG] [ERR] 点击麦克风失败，跳过")
                self.stats['click_failed'] += 1
                self.stats['failed'] += 1
                print("="*60 + "\n")
                return False
            print("[LOG] [4/9] 麦克风已启动，开始录音...")
            
            # 5. TTS朗读
            print("[LOG] [5/9] 正在TTS朗读...")
            self.speak_with_system_tts(word)
            print("[LOG] [5/9] TTS朗读完成")
            
            # 6. 等待录音完成
            print("[LOG] [6/9] 等待录音完成...")
            self.wait_for_recording()
            print("[LOG] [6/9] 等待完成")
            
            # 7. 点击停止按钮
            print("[LOG] [7/9] 正在点击停止按钮...")
            stop_success = self.click_stop_button()
            if not stop_success:
                print("[LOG] [WARN] 点击停止按钮失败")
            print("[LOG] [7/9] 停止按钮点击完成")
            
            # 8. 等待录音完毕后的延迟
            post_delay = self.config.get('recording_settings', {}).get('post_recording_delay', 3.0)
            print(f"[LOG] [8/9] 录音完毕，等待 {post_delay} 秒...")
            time.sleep(post_delay)
            print("[LOG] [8/9] 等待完成")
            
            # 9. 点击继续按钮
            print("[LOG] [9/9] 正在点击继续按钮...")
            continue_success = self.click_continue_button()
            if not continue_success:
                print("[LOG] [WARN] 点击继续按钮失败")
            print("[LOG] [9/9] 继续按钮点击完成")
            
            # 更新统计
            self.stats['success'] += 1
            
            print("\n" + "="*60)
            print(f"[LOG] ✓ 处理完成！成功: {self.stats['success']}/{self.stats['total']}")
            print("="*60 + "\n")
            return True
            
        except Exception as e:
            print(f"[ERR] 处理出错: {e}")
            import traceback
            traceback.print_exc()
            self.stats['failed'] += 1
            return False
    
    def print_statistics(self):
        """显示统计信息"""
        print("\n" + "="*60)
        print("   处理统计")
        print("="*60)
        print(f"  总处理次数: {self.stats['total']}")
        print(f"  成功次数:   {self.stats['success']}")
        print(f"  失败次数:   {self.stats['failed']}")
        if self.stats['total'] > 0:
            success_rate = (self.stats['success'] / self.stats['total']) * 100
            print(f"  成功率:     {success_rate:.1f}%")
        print(f"  OCR失败:    {self.stats['ocr_failed']}")
        print(f"  点击失败:   {self.stats['click_failed']}")
        print("="*60 + "\n")
    
    def run_automation(self, mode="single", interval=5, max_count=None):
        """运行自动化任务"""
        print("\n" + "="*60)
        print("   启动单词自动化工具")
        print("="*60 + "\n")
        
        if mode == "single":
            return self.process_single_word()
        
        elif mode == "continuous":
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
                    print(f"\n{'-'*40}")
                    print(f"  第 {count} 次循环")
                    print(f"{'-'*40}")
                    
                    self.process_single_word()
                    
                    print(f"等待 {interval} 秒...")
                    time.sleep(interval)
                    
            except KeyboardInterrupt:
                print("\n\n用户中断执行")
            
            print(f"\n总共执行了 {count} 次")
            self.print_statistics()


def main():
    # 确保从脚本所在目录读取config.json
    script_dir = os.path.dirname(os.path.abspath(__file__))
    config_path = os.path.join(script_dir, "config.json")
    
    print(f"[MAIN] 脚本目录: {script_dir}")
    print(f"[MAIN] 配置文件路径: {config_path}")
    
    automator = EnhancedWordAutomator(config_file=config_path)
    
    if not automator.connect_device():
        return
    
    print("\n请选择模式：")
    print("  [1] 单次执行")
    print("  [2] 连续循环")
    print("  [3] 校准坐标（测试点击位置）")
    
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
        print("\n校准模式：将连续点击当前设置的麦克风位置")
        print("观察点击位置是否准确，如需修改请编辑 config.json\n")
        
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
