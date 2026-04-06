# -*- coding: utf-8 -*-
"""
Kokoro-82M TTS 测试脚本
用于测试Kokoro是否能正常初始化和工作
"""

import sys
import time

print("=" * 60)
print("   Kokoro-82M TTS 测试")
print("=" * 60)

print("\n[1] 检查依赖...")
try:
    from kokoro import KPipeline
    import soundfile as sf
    import sounddevice as sd
    import numpy as np
    print("[OK] 所有依赖已安装")
except ImportError as e:
    print(f"[ERR] 缺少依赖: {e}")
    print("\n请先安装依赖:")
    print("  pip install kokoro misaki[zh] soundfile sounddevice huggingface-hub")
    sys.exit(1)

print("\n[2] 初始化Kokoro...")
print("    注意: 首次运行需要下载模型，请保持网络连接")
print("    下载过程可能需要几分钟，取决于你的网络速度")
print("    正在初始化...")

try:
    start_time = time.time()
    pipeline = KPipeline(lang_code='a')
    elapsed = time.time() - start_time
    print(f"[OK] Kokoro初始化成功，耗时 {elapsed:.1f} 秒")
except Exception as e:
    print(f"[ERR] Kokoro初始化失败: {e}")
    print("\n可能原因:")
    print("  1. 内存不足（建议4GB+可用内存）")
    print("  2. 网络问题（无法下载模型）")
    print("  3. HuggingFace连接问题")
    sys.exit(1)

print("\n[3] 测试文本生成...")
test_text = "Hello, this is a test of Kokoro TTS."
print(f"    测试文本: {test_text}")

try:
    voice = 'zf_xiaoxiao'
    print(f"    使用音色: {voice}")
    generator = pipeline(test_text, voice=voice)
    result = next(generator)
    print("[OK] 文本生成成功")
except Exception as e:
    print(f"[ERR] 文本生成失败: {e}")
    sys.exit(1)

print("\n[4] 提取音频数据...")
try:
    audio_tensor = result.output.audio
    audio_numpy = audio_tensor.detach().cpu().numpy()
    
    if audio_numpy.ndim == 1:
        audio_numpy = audio_numpy.reshape(-1, 1)
    
    sample_rate = 24000
    print(f"[OK] 音频数据已提取")
    print(f"    采样率: {sample_rate} Hz")
    print(f"    音频形状: {audio_numpy.shape}")
except Exception as e:
    print(f"[ERR] 音频数据提取失败: {e}")
    sys.exit(1)

print("\n[5] 播放音频...")
print("    即将播放测试音频，请确保音量适中")
time.sleep(1)

try:
    sd.play(audio_numpy, sample_rate)
    sd.wait()
    print("[OK] 音频播放完成")
except Exception as e:
    print(f"[ERR] 音频播放失败: {e}")
    print("\n可能原因:")
    print("  1. 没有音频输出设备")
    print("  2. 音频设备被占用")

print("\n" + "=" * 60)
print("   测试完成！")
print("=" * 60)

