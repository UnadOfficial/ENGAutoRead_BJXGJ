from abc import ABC, abstractmethod
import time


class TTSEngine(ABC):

    def __init__(self, config):
        self.config = config
        self.is_available = False

    @abstractmethod
    def initialize(self):
        pass

    @abstractmethod
    def speak(self, text):
        pass

    @abstractmethod
    def cleanup(self):
        pass

    def get_duration(self):
        return self.config.get('tts_duration', 2.5)


class WindowsTTSEngine(TTSEngine):

    def __init__(self, config):
        super().__init__(config)
        self.engine = None

        try:
            import pyttsx3
            self._pyttsx3 = pyttsx3
            self.has_dependency = True
        except ImportError:
            self.has_dependency = False
            print("[WARN] pyttsx3 未安装，Windows TTS 不可用")

    def initialize(self):
        if not self.has_dependency:
            return False

        try:
            self.engine = self._pyttsx3.init()
            tts_cfg = self.config.get('tts_settings', {})
            self.engine.setProperty('rate', int(150 * tts_cfg.get('speed', 1.0)))
            for voice in self.engine.getProperty('voices'):
                if 'english' in voice.name.lower() or 'en' in voice.id.lower():
                    self.engine.setProperty('voice', voice.id)
                    break
            self.is_available = True
            print("[OK] Windows TTS引擎初始化成功")
            return True
        except Exception as e:
            print(f"[WARN] Windows TTS初始化失败: {e}")
            self.is_available = False
            return False

    def speak(self, text):
        if not self.is_available or not self.engine:
            raise Exception("Windows TTS 引擎不可用")

        try:
            print(f"[TTS] 朗读: {text}")
            self.engine.say(text)
            self.engine.runAndWait()
            time.sleep(self.get_duration())
        except Exception as e:
            print(f"[WARN] Windows TTS朗读失败: {e}")
            raise

    def cleanup(self):
        if self.engine:
            try:
                self.engine.stop()
            except Exception as e:
                print(f"[WARN] Windows TTS清理失败: {e}")


class KokoroTTSEngine(TTSEngine):

    def __init__(self, config):
        super().__init__(config)
        self.pipeline = None
        self._initialized = False

        try:
            from kokoro import KPipeline
            import sounddevice as sd
            import numpy as np
            self._KPipeline = KPipeline
            self._sd = sd
            self._np = np
            self.has_dependency = True
        except ImportError:
            self.has_dependency = False
            print("[WARN] Kokoro 依赖未安装，Kokoro TTS 不可用")

    def initialize(self):
        if not self.has_dependency:
            return False
        self.is_available = True
        return True

    def _ensure_initialized(self):
        if self._initialized:
            return True

        try:
            print("[Kokoro] 正在初始化（首次需下载模型约300MB）...")
            self.pipeline = self._KPipeline(lang_code='a')
            self._initialized = True
            self.is_available = True
            print("[OK] Kokoro初始化成功")
            return True
        except Exception as e:
            print(f"[WARN] Kokoro初始化失败: {e}")
            self.is_available = False
            return False

    def speak(self, text):
        if not self.is_available:
            raise Exception("Kokoro TTS 引擎不可用")

        if not self._ensure_initialized():
            raise Exception("Kokoro 初始化失败")

        try:
            voice = self.config.get('tts_settings', {}).get('kokoro_voice', 'zm_yunjian')
            print(f"[Kokoro] 朗读: {text} (音色: {voice})")
            generator = self.pipeline(text, voice=voice)
            result = next(generator)
            audio_numpy = result.output.audio.detach().cpu().numpy()
            if audio_numpy.ndim == 1:
                audio_numpy = audio_numpy.reshape(-1, 1)
            self._sd.play(audio_numpy, 24000)
            self._sd.wait()
            time.sleep(self.get_duration())
        except Exception as e:
            print(f"[WARN] Kokoro朗读失败: {e}")
            raise

    def cleanup(self):
        pass


class TTSEngineFactory:

    @staticmethod
    def create_engine(engine_type, config):
        if engine_type == 'kokoro':
            return KokoroTTSEngine(config)
        elif engine_type == 'windows':
            return WindowsTTSEngine(config)
        else:
            raise ValueError(f"不支持的 TTS 引擎类型: {engine_type}")

    @staticmethod
    def create_fallback_chain(config):
        engines = []
        primary_type = config.get('tts_settings', {}).get('engine', 'windows')
        primary_engine = TTSEngineFactory.create_engine(primary_type, config)
        engines.append(primary_engine)

        if primary_type != 'windows':
            fallback_engine = TTSEngineFactory.create_engine('windows', config)
            engines.append(fallback_engine)

        return engines
