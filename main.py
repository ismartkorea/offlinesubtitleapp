import tkinter as tk
from tkinter import ttk, messagebox
from threading import Thread
import pyaudio
import soundfile as sf
import subprocess
import os
import time
import win32gui, win32con
import signal
import sys
import numpy as np
import wave
import tkinter.colorchooser as colorchooser
import tkinter.font as tkfont

# sounddevice를 선택적 import
try:
    import sounddevice as sd
    SOUNDDEVICE_AVAILABLE = True
    print("✅ sounddevice 라이브러리 사용 가능")
except ImportError:
    SOUNDDEVICE_AVAILABLE = False
    print("⚠️ sounddevice 라이브러리가 없습니다. 출력 장치 캡처 기능이 제한됩니다.")
    print("💡 pip install sounddevice로 설치하면 출력 장치 캡처가 가능합니다.")

from transformers import M2M100ForConditionalGeneration, M2M100Tokenizer

# ========== 설정 ==========
WHISPER_CPP_DIR = os.path.join(os.getcwd(), "whisper.cpp")
WHISPER_MODEL = os.path.join(WHISPER_CPP_DIR, "models", "ggml-base.bin")
WHISPER_EXE = os.path.join(WHISPER_CPP_DIR, "whisper-cli.exe")
RECORD_SECONDS = 1  # 1초마다 녹음
LOOP_DELAY = 1
CHUNK = 1024
FORMAT = pyaudio.paInt16
CHANNELS = 1
RATE = 16000

# 전역 변수로 선택된 장치 저장
selected_device_index = None
selected_device_info = None

# ========== 번역 모델 초기화 ==========
tokenizer = M2M100Tokenizer.from_pretrained("facebook/m2m100_418M")
model = M2M100ForConditionalGeneration.from_pretrained("facebook/m2m100_418M")

def signal_handler(signum, frame):
    """시그널 핸들러 - 프로그램 종료 시 호출"""
    print("\n🛑 프로그램 종료 신호를 받았습니다...")
    sys.exit(0)

# 시그널 핸들러 등록
signal.signal(signal.SIGINT, signal_handler)
signal.signal(signal.SIGTERM, signal_handler)

def detect_language(text):
    if any("\uac00" <= c <= "\ud7a3" for c in text): return "ko"
    elif any("a" <= c.lower() <= "z" for c in text): return "en"
    elif any("\u3040" <= c <= "\u309f" for c in text): return "ja"
    else: return "en"

def translate_text(text, source_lang, target_lang):
    if not text: return ""
    tokenizer.src_lang = source_lang
    encoded = tokenizer(text, return_tensors="pt")
    generated_tokens = model.generate(**encoded, forced_bos_token_id=tokenizer.get_lang_id(target_lang))
    return tokenizer.batch_decode(generated_tokens, skip_special_tokens=True)[0]

def get_audio_devices():
    """사용 가능한 오디오 장치 목록을 가져옵니다"""
    p = pyaudio.PyAudio()
    devices = []
    
    print("=== 전체 오디오 장치 목록 ===")
    for i in range(p.get_device_count()):
        device_info = p.get_device_info_by_index(i)
        print(f"장치 {i}: {device_info['name']} (입력: {device_info['maxInputChannels']}, 출력: {device_info['maxOutputChannels']})")
        
        # 입력 장치 또는 특별한 장치들 포함
        if (device_info['maxInputChannels'] > 0 or 
            'stereo mix' in device_info['name'].lower() or
            'what u hear' in device_info['name'].lower() or
            'loopback' in device_info['name'].lower() or
            'headphones' in device_info['name'].lower() or
            'speakers' in device_info['name'].lower() or
            'capture' in device_info['name'].lower() or
            'webcam' in device_info['name'].lower() or
            'camera' in device_info['name'].lower() or
            'video' in device_info['name'].lower() or
            'hdmi' in device_info['name'].lower() or
            'displayport' in device_info['name'].lower() or
            'usb' in device_info['name'].lower()):
            
            devices.append({
                'index': i,
                'name': device_info['name'],
                'input_channels': device_info['maxInputChannels'],
                'output_channels': device_info['maxOutputChannels'],
                'sample_rate': int(device_info['defaultSampleRate']),
                'is_input': device_info['maxInputChannels'] > 0,
                'is_output': device_info['maxOutputChannels'] > 0,
                'is_video_capture': any(keyword in device_info['name'].lower() 
                                      for keyword in ['capture', 'webcam', 'camera', 'video', 'hdmi', 'displayport'])
            })
    
    p.terminate()
    return devices

class DeviceSelector:
    def __init__(self):
        self.selected_device = None
        self.root = tk.Tk()
        self.root.title("오디오 장치 선택")
        self.root.geometry("600x400")
        self.root.resizable(True, True)
        
        # 중앙 정렬
        self.root.update_idletasks()
        x = (self.root.winfo_screenwidth() // 2) - (600 // 2)
        y = (self.root.winfo_screenheight() // 2) - (400 // 2)
        self.root.geometry(f"600x400+{x}+{y}")
        
        self.setup_ui()
        
    def setup_ui(self):
        # 제목
        title_label = tk.Label(self.root, text="🎵 오디오 캡처 장치 선택", 
                              font=("Arial", 16, "bold"))
        title_label.pack(pady=20)
        
        # 설명
        desc_text = "시스템 오디오를 캡처할 장치를 선택하세요.\n" + \
                   "• 🎵 스테레오 믹스: 데스크탑 오디오 캡처 (권장)\n" + \
                   "• 🎤 마이크로폰: 음성 녹음\n" + \
                   "• 📹 비디오 캡처: 웹캠, 캡처카드 등"
        
        if SOUNDDEVICE_AVAILABLE:
            desc_text += "\n• 🔊 출력 장치: WASAPI Loopback으로 직접 캡처"
        else:
            desc_text += "\n• 🔊 출력 장치: sounddevice 설치 필요"
        
        desc_label = tk.Label(self.root, text=desc_text, 
                             font=("Arial", 10), justify="left")
        desc_label.pack(pady=10)
        
        # 장치 목록 프레임
        list_frame = tk.Frame(self.root)
        list_frame.pack(fill="both", expand=True, padx=20, pady=10)
        
        # 트리뷰 생성
        columns = ("장치명", "입력", "출력", "샘플레이트")
        self.tree = ttk.Treeview(list_frame, columns=columns, show="headings", height=10)
        
        # 컬럼 설정
        self.tree.heading("장치명", text="장치명")
        self.tree.heading("입력", text="입력")
        self.tree.heading("출력", text="출력")
        self.tree.heading("샘플레이트", text="샘플레이트")
        
        self.tree.column("장치명", width=250)
        self.tree.column("입력", width=80)
        self.tree.column("출력", width=80)
        self.tree.column("샘플레이트", width=100)
        
        # 스크롤바
        scrollbar = ttk.Scrollbar(list_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        
        self.tree.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # 장치 목록 로드
        self.load_devices()
        
        # 버튼 프레임
        button_frame = tk.Frame(self.root)
        button_frame.pack(pady=20)
        
        # 선택 버튼
        select_btn = tk.Button(button_frame, text="선택", command=self.select_device,
                              font=("Arial", 12), bg="#4CAF50", fg="white", 
                              width=10, height=2)
        select_btn.pack(side="left", padx=10)
        
        # 새로고침 버튼
        refresh_btn = tk.Button(button_frame, text="새로고침", command=self.refresh_devices,
                               font=("Arial", 12), bg="#2196F3", fg="white",
                               width=10, height=2)
        refresh_btn.pack(side="left", padx=10)
        
        # 취소 버튼
        cancel_btn = tk.Button(button_frame, text="취소", command=self.cancel,
                              font=("Arial", 12), bg="#f44336", fg="white",
                              width=10, height=2)
        cancel_btn.pack(side="left", padx=10)
        
        # 더블클릭 이벤트
        self.tree.bind("<Double-1>", lambda e: self.select_device())
        
    def load_devices(self):
        """오디오 장치 목록을 로드합니다"""
        # 기존 항목 삭제
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        devices = get_audio_devices()
        
        for device in devices:
            # 장치 타입에 따른 아이콘과 태그 설정
            if ('stereo mix' in device['name'].lower() or 
                'what u hear' in device['name'].lower() or
                'loopback' in device['name'].lower()):
                icon = "🎵"
                tags = ('stereo_mix',)
                name_display = f"{icon} {device['name']} (권장)"
            elif device['is_video_capture']:
                icon = "📹"
                tags = ('video_capture',)
                name_display = f"{icon} {device['name']} (비디오 캡처)"
            elif device['is_input'] and device['is_output']:
                icon = "🔄"
                tags = ('both',)
                name_display = f"{icon} {device['name']} (입출력)"
            elif device['is_output']:
                icon = "🔊"
                tags = ('output',)
                name_display = f"{icon} {device['name']} (출력)"
            else:
                icon = "🎤"
                tags = ('input',)
                name_display = f"{icon} {device['name']} (입력)"
            
            self.tree.insert("", "end", values=(
                name_display,
                device['input_channels'],
                device['output_channels'],
                device['sample_rate']
            ), tags=tags)
        
        # 스테레오 믹스 장치가 있으면 첫 번째로 선택
        for item in self.tree.get_children():
            if 'stereo_mix' in self.tree.item(item, "tags"):
                self.tree.selection_set(item)
                break
        else:
            # 스테레오 믹스가 없으면 첫 번째 장치 선택
            if self.tree.get_children():
                self.tree.selection_set(self.tree.get_children()[0])
    
    def refresh_devices(self):
        """장치 목록을 새로고침합니다"""
        self.load_devices()
        messagebox.showinfo("새로고침", "오디오 장치 목록을 새로고침했습니다.")
    
    def select_device(self):
        """선택된 장치를 확인합니다"""
        selection = self.tree.selection()
        if not selection:
            messagebox.showwarning("경고", "장치를 선택해주세요.")
            return
        
        item = selection[0]
        device_name = self.tree.item(item, "values")[0]
        
        # 장치 인덱스 찾기
        devices = get_audio_devices()
        for device in devices:
            if device['name'] in device_name:
                self.selected_device = device
                break
        
        if self.selected_device:
            # 장치 타입에 따른 메시지
            device_type = ""
            if self.selected_device['output_channels'] > 0 and self.selected_device['input_channels'] == 0:
                if SOUNDDEVICE_AVAILABLE:
                    device_type = "출력 장치 (WASAPI Loopback 사용)"
                else:
                    device_type = "출력 장치 (sounddevice 설치 필요)"
            elif self.selected_device['input_channels'] > 0 and self.selected_device['output_channels'] == 0:
                if self.selected_device['is_video_capture']:
                    device_type = "비디오 캡처 장치"
                else:
                    device_type = "입력 장치"
            elif self.selected_device['input_channels'] > 0 and self.selected_device['output_channels'] > 0:
                if self.selected_device['is_video_capture']:
                    device_type = "비디오 캡처 장치 (입출력)"
                else:
                    device_type = "입출력 장치"
            else:
                device_type = "알 수 없는 장치"
            
            # 출력 장치이고 sounddevice가 없는 경우 경고
            warning_msg = ""
            if (self.selected_device['output_channels'] > 0 and 
                self.selected_device['input_channels'] == 0 and 
                not SOUNDDEVICE_AVAILABLE):
                warning_msg = "\n\n⚠️ 출력 장치 캡처를 위해서는 sounddevice 설치가 필요합니다.\npip install sounddevice"
            
            result = messagebox.askyesno("확인", 
                f"선택된 장치: {self.selected_device['name']}\n"
                f"장치 타입: {device_type}\n"
                f"입력 채널: {self.selected_device['input_channels']}\n"
                f"출력 채널: {self.selected_device['output_channels']}\n\n"
                f"이 장치로 오디오를 캡처하시겠습니까?{warning_msg}")
            
            if result:
                global selected_device_index, selected_device_info
                selected_device_index = self.selected_device['index']
                selected_device_info = self.selected_device
                self.root.destroy()
        else:
            messagebox.showerror("오류", "장치 정보를 가져올 수 없습니다.")
    
    def cancel(self):
        """취소하고 프로그램을 종료합니다"""
        if messagebox.askyesno("종료", "프로그램을 종료하시겠습니까?"):
            sys.exit(0)
    
    def run(self):
        """장치 선택 창을 실행합니다"""
        self.root.mainloop()
        return self.selected_device

def capture_audio_with_selected_device(filename="system_audio.wav", duration=RECORD_SECONDS):
    """선택된 장치로 오디오를 캡처합니다"""
    global selected_device_index, selected_device_info
    
    if selected_device_index is None:
        print("❌ 오디오 장치가 선택되지 않았습니다.")
        return None
    
    print(f"🎵 선택된 장치로 오디오 캡처 중... (장치 인덱스: {selected_device_index})")
    print(f"📁 저장할 파일: {filename}")
    print(f"⏱️ 녹음 시간: {duration}초")
    
    # 장치 정보 확인
    if selected_device_info:
        print(f"장치 정보: {selected_device_info['name']}")
        print(f"입력 채널: {selected_device_info['input_channels']}, 출력 채널: {selected_device_info['output_channels']}")
        
        # 출력 장치인 경우 WASAPI Loopback 사용
        if selected_device_info['output_channels'] > 0 and selected_device_info['input_channels'] == 0:
            print("🔄 출력 장치 감지 - WASAPI Loopback 사용")
            result = capture_output_device_audio(filename, duration)
        else:
            print("🔄 입력 장치 감지 - 일반 캡처 사용")
            result = capture_input_device_audio(filename, duration)
    else:
        print("🔄 장치 정보 없음 - 일반 캡처 사용")
        result = capture_input_device_audio(filename, duration)
    
    # 파일 생성 확인
    if result and os.path.exists(filename):
        file_size = os.path.getsize(filename)
        print(f"✅ 오디오 파일 생성 완료: {filename} ({file_size} bytes)")
        return filename
    else:
        print(f"❌ 오디오 파일 생성 실패: {filename}")
        if os.path.exists(filename):
            print(f"⚠️ 파일은 존재하지만 크기가 0일 수 있습니다: {os.path.getsize(filename)} bytes")
        return None

def capture_output_device_audio(filename="system_audio.wav", duration=RECORD_SECONDS):
    """출력 장치에서 WASAPI Loopback으로 오디오를 캡처합니다"""
    global selected_device_index
    
    if not SOUNDDEVICE_AVAILABLE:
        print("❌ sounddevice 라이브러리가 없어서 출력 장치 캡처를 할 수 없습니다.")
        print("💡 pip install sounddevice로 설치해주세요.")
        return None
    
    print("🔊 WASAPI Loopback으로 출력 장치 캡처 중...")
    print(f"📁 저장할 파일: {filename}")
    print(f"⏱️ 녹음 시간: {duration}초")
    
    try:
        print("🔧 sounddevice 스트림 생성 중...")
        # sounddevice를 사용하여 출력 장치에서 오디오 캡처
        # WASAPI Loopback 모드로 출력 장치의 오디오를 캡처
        audio_data = sd.rec(
            int(duration * RATE),
            samplerate=RATE,
            channels=CHANNELS,
            dtype='float32',
            device=selected_device_index,  # 선택된 출력 장치
            latency='low'
        )
        
        print("🎙️ 오디오 데이터 수집 중...")
        # 녹음 완료까지 대기
        sd.wait()
        
        print(f"📦 수집된 오디오 데이터 크기: {audio_data.shape}")
        
        if audio_data.size == 0:
            print("❌ 수집된 오디오 데이터가 없습니다.")
            return None
        
        print("💾 WAV 파일로 저장 중...")
        # 오디오 데이터를 WAV 파일로 저장
        sf.write(filename, audio_data, RATE)
        
        print("✅ 출력 장치 오디오 캡처 완료")
        return filename
        
    except Exception as e:
        print(f"❌ 출력 장치 오디오 캡처 실패: {e}")
        import traceback
        traceback.print_exc()
        print("🔄 일반 입력 장치로 대체 시도...")
        return capture_input_device_audio(filename, duration)

def capture_input_device_audio(filename="system_audio.wav", duration=RECORD_SECONDS):
    """일반 입력 장치로 오디오를 캡처합니다"""
    global selected_device_index
    
    print("🎤 일반 입력 장치로 오디오 캡처 중...")
    
    p = pyaudio.PyAudio()
    
    try:
        # 장치 정보 확인
        device_info = p.get_device_info_by_index(selected_device_index)
        print(f"장치 정보: {device_info['name']}")
        print(f"입력 채널: {device_info['maxInputChannels']}, 출력 채널: {device_info['maxOutputChannels']}")
        
        # 입력 채널이 있는지 확인
        if device_info['maxInputChannels'] == 0:
            print("⚠️ 선택된 장치는 입력 장치가 아닙니다.")
            print("💡 다른 장치를 선택하거나 스테레오 믹스를 활성화해주세요.")
            p.terminate()
            return None
        
        print("🔧 PyAudio 스트림 생성 중...")
        # 선택된 장치로 오디오 캡처
        stream = p.open(format=FORMAT,
                        channels=CHANNELS,
                        rate=RATE,
                        input=True,
                        input_device_index=selected_device_index,
                        frames_per_buffer=CHUNK)
        
        print("🎙️ 오디오 데이터 수집 시작...")
        frames = []
        total_chunks = int(RATE / CHUNK * duration)
        
        # 오디오 데이터 수집
        for i in range(0, total_chunks):
            try:
                data = stream.read(CHUNK)
                frames.append(data)
                if i % 10 == 0:  # 10개 청크마다 진행상황 출력
                    progress = (i / total_chunks) * 100
                    print(f"📊 녹음 진행률: {progress:.1f}% ({i}/{total_chunks})")
            except Exception as e:
                print(f"❌ 청크 {i} 읽기 실패: {e}")
                break
        
        print(f"📦 수집된 프레임 수: {len(frames)}")
        
        if len(frames) == 0:
            print("❌ 수집된 오디오 데이터가 없습니다.")
            stream.stop_stream()
            stream.close()
            p.terminate()
            return None
        
        print("✅ 입력 장치 오디오 캡처 완료")
        
        # 스트림 정리
        stream.stop_stream()
        stream.close()
        p.terminate()
        
        print("💾 WAV 파일로 저장 중...")
        # WAV 파일로 저장
        wf = wave.open(filename, 'wb')
        wf.setnchannels(CHANNELS)
        wf.setsampwidth(p.get_sample_size(FORMAT))
        wf.setframerate(RATE)
        wf.writeframes(b''.join(frames))
        wf.close()
        
        print(f"✅ WAV 파일 저장 완료: {filename}")
        return filename
        
    except Exception as e:
        print(f"❌ 입력 장치 오디오 캡처 실패: {e}")
        import traceback
        traceback.print_exc()
        p.terminate()
        return None

def run_whisper_cpp(audio_path="system_audio.wav"):
    print(f"🔍 Whisper 실행 중: {WHISPER_EXE}")
    print(f"📁 오디오 파일: {audio_path}")
    print(f"📁 모델 파일: {WHISPER_MODEL}")
    
    if not os.path.exists(WHISPER_EXE):
        print(f"❌ Whisper 실행 파일을 찾을 수 없습니다: {WHISPER_EXE}")
        return ""
    
    if not os.path.exists(WHISPER_MODEL):
        print(f"❌ Whisper 모델 파일을 찾을 수 없습니다: {WHISPER_MODEL}")
        return ""
    
    if not os.path.exists(audio_path):
        print(f"❌ 오디오 파일을 찾을 수 없습니다: {audio_path}")
        return ""
    
    try:
        result = subprocess.run([
            WHISPER_EXE,
            "--model", WHISPER_MODEL,
            "--file", audio_path,
            "--output-txt",
            "--output-file", "result"
        ], capture_output=True, text=True, timeout=30)
        
        print(f"✅ Whisper 실행 완료")
        print(f"📤 출력: {result.stdout}")
        if result.stderr:
            print(f"⚠️ 오류: {result.stderr}")
        
        result_path = "result.txt"
        if os.path.exists(result_path):
            with open(result_path, "r", encoding="utf-8") as f:
                text = f.read().strip()
                print(f"📝 인식된 텍스트: {text}")
                return text
        else:
            print(f"❌ 결과 파일을 찾을 수 없습니다: {result_path}")
            return ""
    except subprocess.TimeoutExpired:
        print("⏰ Whisper 실행 시간 초과")
        return ""
    except Exception as e:
        print(f"❌ Whisper 실행 중 오류: {e}")
        return ""

def speech_loop(update_fn, app_instance):
    print("🎬 실시간 자막 루프 시작")
    while app_instance.running:
        try:
            capture_audio_with_selected_device(duration=RECORD_SECONDS)
            text = run_whisper_cpp()
            
            if not app_instance.running:  # 종료 신호 확인
                break
                
            if text:
                src_lang = detect_language(text)
                tgt_lang = "en" if src_lang != "en" else "ko"
                translated = translate_text(text, src_lang, tgt_lang)
                display = f"{translated}"
                print(f"🌐 번역 결과: {display}")
            else:
                display = "🎧 음성을 인식하지 못했습니다..."
                print("🔇 음성 인식 실패")
            
            update_fn(display)
            # sleep을 0.1~0.5초로 줄이면 더 빠름
            time.sleep(0.1)
        except Exception as e:
            print(f"❌ 루프 실행 중 오류: {e}")
            update_fn("⚠️ 오류가 발생했습니다...")
            time.sleep(0.5)
    
    print("🛑 음성 인식 루프 종료")

def make_window_clickthrough(hwnd):
    styles = win32gui.GetWindowLong(hwnd, win32con.GWL_EXSTYLE)
    styles |= win32con.WS_EX_LAYERED | win32con.WS_EX_TRANSPARENT | win32con.WS_EX_TOPMOST
    win32gui.SetWindowLong(hwnd, win32con.GWL_EXSTYLE, styles)

class OverlaySubtitleApp:
    def __init__(self, root):
        self.root = root
        self.running = True  # 인스턴스 변수로 종료 플래그 관리
        
        self.root.overrideredirect(True)
        self.root.attributes('-topmost', True)
        self.root.attributes('-alpha', 0.6)
        self.root.configure(bg='black')

        # 창 크기 기본값
        self._width = 600
        self._height = 80
        self.root.geometry(f"{self._width}x{self._height}+100+100")

        # 기본 스타일
        self.bg_color = 'black'
        self.fg_color = 'white'
        self.font_family = 'Arial'
        self.font_size = 28
        self.label = tk.Label(root, text="🎧 자막 준비 중...", font=(self.font_family, self.font_size),
                              fg=self.fg_color, bg=self.bg_color, wraplength=self._width, justify="center")
        self.label.pack(expand=True, fill="both")

        # 설정 버튼
        self.settings_btn = tk.Button(root, text="⚙️", command=self.open_settings, font=("Arial", 14), bg="#333", fg="white", bd=0)
        self.settings_btn.place(relx=1.0, rely=0.0, anchor="ne", x=-10, y=10)

        # 종료 버튼
        self.close_btn = tk.Button(root, text="❌", command=self.on_closing, font=("Arial", 14), bg="#333", fg="white", bd=0)
        self.close_btn.place(relx=1.0, rely=0.0, anchor="ne", x=-50, y=10)

        # 창 이동 관련 변수
        self._offset_x = 0
        self._offset_y = 0
        self.root.bind('<ButtonPress-1>', self.start_move)
        self.root.bind('<B1-Motion>', self.do_move)

        # 창 크기 조절 관련 변수
        self._resizing = False
        self._resize_start_x = 0
        self._resize_start_y = 0
        self._resize_start_width = self._width
        self._resize_start_height = self._height
        self.root.bind('<ButtonPress-3>', self.start_resize)  # 우클릭으로 시작
        self.root.bind('<B3-Motion>', self.do_resize)
        self.root.bind('<ButtonRelease-3>', self.stop_resize)

        # 종료 이벤트 바인딩
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
        self.root.bind('<Escape>', self.on_closing)
        self.root.bind('<Control-c>', self.on_closing)

        self.thread = Thread(target=speech_loop, args=(self.update_text, self))
        self.thread.daemon = True
        self.thread.start()

    def start_move(self, event):
        self._offset_x = event.x
        self._offset_y = event.y

    def do_move(self, event):
        if not self._resizing:
            x = self.root.winfo_x() + event.x - self._offset_x
            y = self.root.winfo_y() + event.y - self._offset_y
            self.root.geometry(f"+{x}+{y}")

    def start_resize(self, event):
        self._resizing = True
        self._resize_start_x = event.x_root
        self._resize_start_y = event.y_root
        self._resize_start_width = self.root.winfo_width()
        self._resize_start_height = self.root.winfo_height()

    def do_resize(self, event):
        if self._resizing:
            dx = event.x_root - self._resize_start_x
            dy = event.y_root - self._resize_start_y
            new_width = max(200, self._resize_start_width + dx)
            new_height = max(50, self._resize_start_height + dy)
            self.root.geometry(f"{new_width}x{new_height}")
            self.label.config(wraplength=new_width)

    def stop_resize(self, event):
        self._resizing = False

    def update_text(self, text):
        if hasattr(self, 'label') and self.label.winfo_exists():
            self.label.config(text=text)

    def open_settings(self):
        settings_win = tk.Toplevel(self.root)
        settings_win.title("자막 설정")
        settings_win.geometry("320x250")
        settings_win.resizable(False, False)
        settings_win.attributes('-topmost', True)

        # 배경색
        tk.Label(settings_win, text="배경색:").pack(pady=(20, 0))
        bg_btn = tk.Button(settings_win, text="배경색 선택", command=self.choose_bg_color)
        bg_btn.pack(pady=5)

        # 글자색
        tk.Label(settings_win, text="글자색:").pack(pady=(10, 0))
        fg_btn = tk.Button(settings_win, text="글자색 선택", command=self.choose_fg_color)
        fg_btn.pack(pady=5)

        # 폰트
        tk.Label(settings_win, text="글꼴/크기:").pack(pady=(10, 0))
        font_btn = tk.Button(settings_win, text="글꼴/크기 선택", command=self.choose_font)
        font_btn.pack(pady=5)

        # 닫기
        close_btn = tk.Button(settings_win, text="닫기", command=settings_win.destroy)
        close_btn.pack(pady=20)

    def choose_bg_color(self):
        color = colorchooser.askcolor(title="배경색 선택", initialcolor=self.bg_color)[1]
        if color:
            self.bg_color = color
            self.label.config(bg=self.bg_color)
            self.root.configure(bg=self.bg_color)

    def choose_fg_color(self):
        color = colorchooser.askcolor(title="글자색 선택", initialcolor=self.fg_color)[1]
        if color:
            self.fg_color = color
            self.label.config(fg=self.fg_color)

    def choose_font(self):
        # 폰트 선택 다이얼로그 (tkinter 기본은 없음, 간단 구현)
        font_win = tk.Toplevel(self.root)
        font_win.title("글꼴/크기 선택")
        font_win.geometry("300x180")
        font_win.attributes('-topmost', True)
        tk.Label(font_win, text="글꼴명:").pack(pady=(10, 0))
        font_entry = tk.Entry(font_win)
        font_entry.insert(0, self.font_family)
        font_entry.pack(pady=5)
        tk.Label(font_win, text="크기:").pack(pady=(10, 0))
        size_entry = tk.Entry(font_win)
        size_entry.insert(0, str(self.font_size))
        size_entry.pack(pady=5)
        def apply_font():
            family = font_entry.get()
            try:
                size = int(size_entry.get())
            except ValueError:
                size = self.font_size
            self.font_family = family
            self.font_size = size
            self.label.config(font=(self.font_family, self.font_size))
            font_win.destroy()
        apply_btn = tk.Button(font_win, text="적용", command=apply_font)
        apply_btn.pack(pady=10)

    def on_closing(self, event=None):
        """프로그램 종료 처리"""
        print("🛑 프로그램 종료 요청...")
        self.running = False
        if hasattr(self, 'root') and self.root.winfo_exists():
            self.root.quit()
            self.root.destroy()
        sys.exit(0)

    def cleanup(self):
        """정리 작업"""
        self.running = False
        if hasattr(self, 'root') and self.root.winfo_exists():
            self.root.quit()
            self.root.destroy()

if __name__ == "__main__":
    print("🎵 오프라인 자막 앱 시작...")
    
    # 장치 선택 창 표시
    device_selector = DeviceSelector()
    selected_device = device_selector.run()
    
    if selected_device is None:
        print("❌ 장치가 선택되지 않았습니다. 프로그램을 종료합니다.")
        sys.exit(0)
    
    print(f"✅ 선택된 장치: {selected_device['name']}")
    
    # 메인 자막 앱 시작
    app = None
    try:
        root = tk.Tk()
        root.title("OfflineTranslatorOverlay")
        app = OverlaySubtitleApp(root)
        
        # 종료 시그널 처리
        def on_exit():
            if app:
                app.cleanup()
        
        # Ctrl+C 처리
        import atexit
        atexit.register(on_exit)
        
        root.mainloop()
    except KeyboardInterrupt:
        print("\n🛑 Ctrl+C로 프로그램 종료...")
        if app:
            app.cleanup()
    except Exception as e:
        print(f"❌ 프로그램 실행 중 오류: {e}")
        if app:
            app.cleanup()
    finally:
        print("🛑 프로그램 종료 완료")
