import { useState, useEffect } from 'react';
import axios from 'axios';

interface SubtitleItem {
  text: string;
  timestamp: number;
}

export default function App() {
  const [isRecording, setIsRecording] = useState(false);
  const [subtitles, setSubtitles] = useState<SubtitleItem[]>([]);
  const [audioUrl, setAudioUrl] = useState<string>('');

  const startRecognition = () => {
    const recognition = new (window as any).webkitSpeechRecognition();
    recognition.continuous = true;
    recognition.interimResults = true;

    recognition.onresult = (event: any) => {
      const transcript = event.results[event.results.length-1][0].transcript;
      setSubtitles(prev => [...prev, {
        text: transcript,
        timestamp: Date.now()
      }]);
    };

    recognition.start();
    setIsRecording(true);
  };

  const generateDubbing = async () => {
    try {
      const response = await axios.post('/api/synthesize', {
        text: subtitles.map(s => s.text).join(' ')
      });
      setAudioUrl(response.data.audioUrl);
    } catch (error) {
      console.error('配音生成失败:', error);
    }
  };

  return (
    <div className="app-container">
      <div className="control-panel">
        <button 
          onClick={isRecording ? () => setIsRecording(false) : startRecognition}
          className={isRecording ? 'recording' : ''}
        >
          {isRecording ? '停止录音' : '开始录音'}
        </button>
        <button onClick={generateDubbing} disabled={!subtitles.length}>
          生成配音
        </button>
      </div>

      <div className="subtitle-display">
        {subtitles.map((sub, index) => (
          <div key={index} className="subtitle-item">
            {sub.text}
          </div>
        ))}
      </div>

      {audioUrl && (
        <audio controls src={audioUrl} className="audio-player" />
      )}
    </div>
  );
}