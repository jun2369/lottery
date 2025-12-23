import React, { useState, useEffect, useCallback, useRef } from 'react';
import './App.css';

// 声明全局 XLSX 类型
declare global {
  interface Window {
    XLSX: any;
  }
}

// Types
interface Participant {
  name: string;
  won: boolean;
  prize?: string;
}

interface Winner {
  name: string;
  prize: string;
  time: string;
}

interface StorageData {
  participants: Participant[];
  winners: Winner[];
  prizeConfig?: PrizeConfig[];
}

interface PrizeConfig {
  name: string;
  label: string;
  count: number;
  drawn: number;
}

// Color scheme
const COLORS = [
  '#FF6B6B', '#4ECDC4', '#45B7D1', '#96CEB4', '#FFEAA7',
  '#DDA0DD', '#98D8C8', '#F7DC6F', '#BB8FCE', '#85C1E9',
  '#F8B500', '#00CED1', '#FF69B4', '#32CD32', '#FF8C00',
  '#9370DB', '#20B2AA', '#FFD700', '#FF6347', '#40E0D0'
];

const DEFAULT_PRIZE_CONFIG: PrizeConfig[] = [
  { name: 'Grand Prize', label: '🏆 Grand Prize', count: 1, drawn: 0 },
  { name: '1st Prize', label: '🥇 1st Prize', count: 2, drawn: 0 },
  { name: '2nd Prize', label: '🥈 2nd Prize', count: 3, drawn: 0 },
  { name: '3rd Prize', label: '🥉 3rd Prize', count: 5, drawn: 0 },
  { name: 'Lucky Prize', label: '🎁 Lucky Prize', count: 10, drawn: 0 },
];

export default function App() {
  const [participants, setParticipants] = useState<Participant[]>([]);
  const [winners, setWinners] = useState<Winner[]>([]);
  const [nameInput, setNameInput] = useState('');
  const [batchInput, setBatchInput] = useState('');
  const [prizeConfig, setPrizeConfig] = useState<PrizeConfig[]>(DEFAULT_PRIZE_CONFIG);
  const [configSaved, setConfigSaved] = useState(true);
  const [selectedPrizeIndex, setSelectedPrizeIndex] = useState(0);
  
  // Sound settings
  const [soundEnabled, setSoundEnabled] = useState(true);
  const [soundVolume, setSoundVolume] = useState(0.7);
  const [voiceEnabled, setVoiceEnabled] = useState(true);
  const [selectedVoice, setSelectedVoice] = useState<SpeechSynthesisVoice | null>(null);
  const [availableVoices, setAvailableVoices] = useState<SpeechSynthesisVoice[]>([]);
  const audioContextRef = useRef<AudioContext | null>(null);
  const tickIntervalRef = useRef<ReturnType<typeof setTimeout> | null>(null);
  const [isSpinning, setIsSpinning] = useState(false);
  const [rotation, setRotation] = useState(0);
  const [spinDuration, setSpinDuration] = useState(8);
  const [currentWinner, setCurrentWinner] = useState<{ name: string; prize: string } | null>(null);
  const [showFireworks, setShowFireworks] = useState(false);
  const [isLoading, setIsLoading] = useState(true);
  const [leftPanelCollapsed, setLeftPanelCollapsed] = useState(false);
  const [rightPanelCollapsed, setRightPanelCollapsed] = useState(false);

  // Load XLSX library
  useEffect(() => {
    if (window.XLSX) {
      return;
    }
    const script = document.createElement('script');
    script.src = 'https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js';
    script.onerror = () => console.error('Failed to load XLSX library');
    document.head.appendChild(script);
  }, []);

  // Initialize Audio Context
  const getAudioContext = () => {
    if (!audioContextRef.current) {
      audioContextRef.current = new (window.AudioContext || (window as any).webkitAudioContext)();
    }
    return audioContextRef.current;
  };

  // Load available voices
  useEffect(() => {
    const loadVoices = () => {
      const voices = speechSynthesis.getVoices();
      if (voices.length > 0) {
        setAvailableVoices(voices);
        // Prefer English voices
        const englishVoice = voices.find(v => v.lang.startsWith('en') && v.name.toLowerCase().includes('female')) ||
                            voices.find(v => v.lang.startsWith('en-US')) ||
                            voices.find(v => v.lang.startsWith('en')) ||
                            voices[0];
        setSelectedVoice(englishVoice || null);
      }
    };
    
    loadVoices();
    speechSynthesis.onvoiceschanged = loadVoices;
    
    return () => {
      speechSynthesis.onvoiceschanged = null;
    };
  }, []);

  // Speak function with emotion control
  const speak = (text: string, rate = 1, pitch = 1, volume = 1) => {
    if (!voiceEnabled) return;
    
    const utterance = new SpeechSynthesisUtterance(text);
    utterance.rate = rate;
    utterance.pitch = pitch;
    utterance.volume = volume;
    if (selectedVoice) {
      utterance.voice = selectedVoice;
    }
    
    speechSynthesis.speak(utterance);
  };

  // Play tick sound (casino wheel click)
  const playTick = (volume = 1, pitch = 800) => {
    if (!soundEnabled) return;
    
    const ctx = getAudioContext();
    const oscillator = ctx.createOscillator();
    const gainNode = ctx.createGain();
    
    oscillator.connect(gainNode);
    gainNode.connect(ctx.destination);
    
    oscillator.frequency.value = pitch;
    oscillator.type = 'square';
    
    const vol = soundVolume * volume * 0.3;
    gainNode.gain.setValueAtTime(vol, ctx.currentTime);
    gainNode.gain.exponentialRampToValueAtTime(0.001, ctx.currentTime + 0.05);
    
    oscillator.start(ctx.currentTime);
    oscillator.stop(ctx.currentTime + 0.05);
  };

  // Play drum roll (building tension)
  const playDrumRoll = (duration: number) => {
    if (!soundEnabled) return;
    
    let interval = 50; // Start fast
    let elapsed = 0;
    
    const tick = () => {
      if (elapsed >= duration * 1000) {
        if (tickIntervalRef.current) {
          clearTimeout(tickIntervalRef.current);
          tickIntervalRef.current = null;
        }
        return;
      }
      
      // Progress from 0 to 1
      const progress = elapsed / (duration * 1000);
      
      // Pitch goes from high to low as it slows
      const pitch = 1200 - (progress * 600);
      
      // Volume increases toward the end
      const volume = 0.5 + (progress * 0.5);
      
      playTick(volume, pitch);
      
      // Interval increases (slows down) as progress increases
      // Exponential slowdown for realistic wheel effect
      interval = 50 + Math.pow(progress, 2) * 400;
      
      elapsed += interval;
      tickIntervalRef.current = setTimeout(tick, interval);
    };
    
    tick();
  };

  // Play fanfare/winner sound (celebration chimes)
  const playFanfare = () => {
    if (!soundEnabled) return;
    
    const ctx = getAudioContext();
    const vol = soundVolume * 0.3;
    
    // Play ascending celebratory chimes
    const notes = [523.25, 659.25, 783.99, 1046.50, 1318.51]; // C5, E5, G5, C6, E6
    
    notes.forEach((freq, i) => {
      const oscillator = ctx.createOscillator();
      const gainNode = ctx.createGain();
      
      oscillator.connect(gainNode);
      gainNode.connect(ctx.destination);
      
      oscillator.frequency.value = freq;
      oscillator.type = 'sine';
      
      const startTime = ctx.currentTime + (i * 0.12);
      gainNode.gain.setValueAtTime(0, startTime);
      gainNode.gain.linearRampToValueAtTime(vol, startTime + 0.02);
      gainNode.gain.exponentialRampToValueAtTime(0.001, startTime + 0.5);
      
      oscillator.start(startTime);
      oscillator.stop(startTime + 0.5);
    });
  };

  // Play big win celebration (slot machine jackpot style)
  const playBigWin = () => {
    if (!soundEnabled) return;
    
    const ctx = getAudioContext();
    const vol = soundVolume * 0.25;
    
    // Jackpot bell sequence - multiple rapid chimes
    const bellFreqs = [1200, 1400, 1600, 1800, 2000, 1800, 1600, 1400, 1200];
    
    bellFreqs.forEach((freq, i) => {
      const oscillator = ctx.createOscillator();
      const gainNode = ctx.createGain();
      
      oscillator.connect(gainNode);
      gainNode.connect(ctx.destination);
      
      oscillator.frequency.value = freq;
      oscillator.type = 'sine';
      
      const startTime = ctx.currentTime + (i * 0.08);
      gainNode.gain.setValueAtTime(vol, startTime);
      gainNode.gain.exponentialRampToValueAtTime(0.001, startTime + 0.2);
      
      oscillator.start(startTime);
      oscillator.stop(startTime + 0.2);
    });
    
    // Add shimmer/sparkle effect
    setTimeout(() => {
      const sparkleFreqs = [2400, 2600, 2800, 3000, 2800, 2600];
      sparkleFreqs.forEach((freq, i) => {
        const osc = ctx.createOscillator();
        const gain = ctx.createGain();
        
        osc.connect(gain);
        gain.connect(ctx.destination);
        
        osc.frequency.value = freq;
        osc.type = 'sine';
        
        const t = ctx.currentTime + (i * 0.05);
        gain.gain.setValueAtTime(vol * 0.5, t);
        gain.gain.exponentialRampToValueAtTime(0.001, t + 0.15);
        
        osc.start(t);
        osc.stop(t + 0.15);
      });
    }, 400);
    
    // Final triumphant chord
    setTimeout(() => playFanfare(), 700);
  };

  // Play start sound (anticipation)
  const playStartSound = () => {
    if (!soundEnabled) return;
    
    const ctx = getAudioContext();
    const vol = soundVolume * 0.3;
    
    // Rising tone
    const oscillator = ctx.createOscillator();
    const gainNode = ctx.createGain();
    
    oscillator.connect(gainNode);
    gainNode.connect(ctx.destination);
    
    oscillator.type = 'sine';
    oscillator.frequency.setValueAtTime(300, ctx.currentTime);
    oscillator.frequency.exponentialRampToValueAtTime(600, ctx.currentTime + 0.3);
    
    gainNode.gain.setValueAtTime(vol, ctx.currentTime);
    gainNode.gain.exponentialRampToValueAtTime(0.001, ctx.currentTime + 0.3);
    
    oscillator.start(ctx.currentTime);
    oscillator.stop(ctx.currentTime + 0.3);
  };

  // Play final countdown beeps
  const playCountdown = () => {
    if (!soundEnabled) return;
    
    const ctx = getAudioContext();
    const vol = soundVolume * 0.5;
    
    // Three beeps getting louder
    [0, 0.3, 0.6].forEach((delay, i) => {
      const oscillator = ctx.createOscillator();
      const gainNode = ctx.createGain();
      
      oscillator.connect(gainNode);
      gainNode.connect(ctx.destination);
      
      oscillator.type = 'sine';
      oscillator.frequency.value = 880; // A5
      
      const startTime = ctx.currentTime + delay;
      const volume = vol * (0.5 + (i * 0.25));
      
      gainNode.gain.setValueAtTime(volume, startTime);
      gainNode.gain.exponentialRampToValueAtTime(0.001, startTime + 0.15);
      
      oscillator.start(startTime);
      oscillator.stop(startTime + 0.15);
    });
  };

  // Cleanup on unmount
  useEffect(() => {
    return () => {
      if (tickIntervalRef.current) {
        clearTimeout(tickIntervalRef.current);
      }
      if (audioContextRef.current) {
        audioContextRef.current.close();
      }
      speechSynthesis.cancel();
    };
  }, []);

  // Load data
  useEffect(() => {
    const loadData = async () => {
      try {
        const saved = localStorage.getItem('lottery_data');
        if (saved) {
          const data: StorageData = JSON.parse(saved);
          setParticipants(data.participants || []);
          setWinners(data.winners || []);
          if (data.prizeConfig) {
            setPrizeConfig(data.prizeConfig);
          }
        }
      } catch (e) {
        console.log('No saved data found');
      } finally {
        setIsLoading(false);
      }
    };
    loadData();
  }, []);

  // Save data
  const saveData = useCallback((p: Participant[], w: Winner[], pc?: PrizeConfig[]) => {
    try {
      const data: StorageData = { 
        participants: p, 
        winners: w,
        prizeConfig: pc || prizeConfig
      };
      localStorage.setItem('lottery_data', JSON.stringify(data));
    } catch (e) {
      console.error('Failed to save data:', e);
    }
  }, [prizeConfig]);

  // Add participant
  const addParticipant = () => {
    const name = nameInput.trim();
    if (!name) return;
    if (participants.find(p => p.name === name)) {
      alert('This name already exists!');
      return;
    }
    const newParticipants = [...participants, { name, won: false }];
    setParticipants(newParticipants);
    setNameInput('');
    saveData(newParticipants, winners);
  };

  // Batch add
  const batchAdd = () => {
    const text = batchInput.trim();
    if (!text) return;
    
    const names = text.split(/[\n,，\s]+/).filter(n => n.trim());
    let added = 0;
    const newParticipants = [...participants];
    
    names.forEach(n => {
      const name = n.trim();
      if (name && !newParticipants.find(p => p.name === name)) {
        newParticipants.push({ name, won: false });
        added++;
      }
    });
    
    setParticipants(newParticipants);
    setBatchInput('');
    saveData(newParticipants, winners);
    alert(`Successfully added ${added} people, ${names.length - added} duplicates skipped`);
  };

  // Download Excel template
  const downloadTemplate = () => {
    if (!window.XLSX) {
      alert('Excel feature is loading, please try again later');
      return;
    }
    const XLSX = window.XLSX;
    const templateData = [
      { 'No.': 1, 'Name': 'John' },
      { 'No.': 2, 'Name': 'Jane' },
      { 'No.': 3, 'Name': 'Bob' },
    ];
    
    const ws = XLSX.utils.json_to_sheet(templateData);
    ws['!cols'] = [{ wch: 8 }, { wch: 20 }];
    
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Participants');
    XLSX.writeFile(wb, 'lottery_template.xlsx');
  };

  // Upload Excel file
  const handleExcelUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    if (!window.XLSX) {
      alert('Excel feature is loading, please try again later');
      return;
    }
    const XLSX = window.XLSX;

    const reader = new FileReader();
    reader.onload = (event) => {
      try {
        const data = new Uint8Array(event.target?.result as ArrayBuffer);
        const workbook = XLSX.read(data, { type: 'array' });
        
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 }) as any[][];
        
        if (jsonData.length < 2) {
          alert('Excel file is empty or has incorrect format!');
          return;
        }

        const headerRow = jsonData[0].map((h: any) => String(h).toLowerCase().trim());
        let nameColIndex = headerRow.findIndex((h: string) => 
          h === 'name' || h === '姓名' || h === '名字' || h === '名称'
        );
        
        if (nameColIndex === -1) {
          nameColIndex = 1;
        }

        const names: string[] = [];
        for (let i = 1; i < jsonData.length; i++) {
          const row = jsonData[i];
          if (row && row[nameColIndex]) {
            const name = String(row[nameColIndex]).trim();
            if (name && name.length > 0) {
              names.push(name);
            }
          }
        }

        if (names.length === 0) {
          alert('Could not read valid names from Excel!');
          return;
        }

        let added = 0;
        let skipped = 0;
        const newParticipants = [...participants];
        
        names.forEach(name => {
          if (!newParticipants.find(p => p.name === name)) {
            newParticipants.push({ name, won: false });
            added++;
          } else {
            skipped++;
          }
        });

        setParticipants(newParticipants);
        saveData(newParticipants, winners);
        alert(`Excel import complete!\nAdded: ${added}\nSkipped (duplicates): ${skipped}`);
      } catch (err) {
        console.error('Excel parse error:', err);
        alert('Failed to parse Excel file, please check the format!');
      }
    };
    reader.readAsArrayBuffer(file);
    e.target.value = '';
  };

  // Remove participant
  const removeParticipant = (index: number) => {
    const newParticipants = participants.filter((_, i) => i !== index);
    setParticipants(newParticipants);
    saveData(newParticipants, winners);
  };

  // Clear all
  const clearAll = () => {
    if (window.confirm('Are you sure you want to clear all participants?')) {
      const newPrizeConfig = prizeConfig.map(pc => ({ ...pc, drawn: 0 }));
      setParticipants([]);
      setWinners([]);
      setPrizeConfig(newPrizeConfig);
      setSelectedPrizeIndex(0);
      saveData([], [], newPrizeConfig);
    }
  };

  // Reset winner status
  const resetWinners = () => {
    if (window.confirm('Are you sure you want to reset all winner status?')) {
      // Stop any ongoing sounds and speech
      if (tickIntervalRef.current) {
        clearTimeout(tickIntervalRef.current);
        tickIntervalRef.current = null;
      }
      speechSynthesis.cancel();
      
      const newParticipants = participants.map(p => ({ ...p, won: false, prize: undefined }));
      const newPrizeConfig = prizeConfig.map(pc => ({ ...pc, drawn: 0 }));
      setParticipants(newParticipants);
      setWinners([]);
      setPrizeConfig(newPrizeConfig);
      setSelectedPrizeIndex(0);
      saveData(newParticipants, [], newPrizeConfig);
    }
  };

  // Export winners as TXT
  const exportWinners = () => {
    if (winners.length === 0) {
      alert('No winner records yet!');
      return;
    }
    let text = '🏆 Winner List 🏆\n================\n\n';
    winners.forEach((w, i) => {
      text += `${i + 1}. ${w.prize}: ${w.name}\n   Time: ${w.time}\n\n`;
    });
    const blob = new Blob([text], { type: 'text/plain' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `winners_${new Date().toLocaleDateString()}.txt`;
    a.click();
    URL.revokeObjectURL(url);
  };

  // Export winners as Excel
  const exportWinnersExcel = () => {
    if (winners.length === 0) {
      alert('No winner records yet!');
      return;
    }
    
    if (!window.XLSX) {
      alert('Excel feature is loading, please try again later');
      return;
    }
    const XLSX = window.XLSX;
    
    const excelData = winners.map((w, index) => ({
      'No.': index + 1,
      'Prize': w.prize,
      'Winner': w.name,
      'Time': w.time
    }));
    
    const ws = XLSX.utils.json_to_sheet(excelData);
    ws['!cols'] = [{ wch: 6 }, { wch: 12 }, { wch: 15 }, { wch: 20 }];
    
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Winners');
    XLSX.writeFile(wb, `winners_${new Date().toLocaleDateString()}.xlsx`);
  };

  // Get selected prize
  const getSelectedPrize = (): PrizeConfig | null => {
    const prize = prizeConfig[selectedPrizeIndex];
    if (prize && prize.count > 0 && prize.drawn < prize.count) {
      return prize;
    }
    return null;
  };

  // Spin the wheel
  const spin = () => {
    const available = participants.filter(p => !p.won);
    if (available.length === 0) {
      alert('No participants available for the draw!');
      return;
    }
    if (isSpinning) return;

    // Get selected prize
    const selectedPrize = getSelectedPrize();
    if (!selectedPrize) {
      alert('Please select an available prize!');
      return;
    }

    // Cancel any ongoing speech
    speechSynthesis.cancel();

    setIsSpinning(true);
    setCurrentWinner(null);

    // Select winner
    const winnerIndex = Math.floor(Math.random() * available.length);
    const winner = available[winnerIndex];
    const actualIndex = participants.findIndex(p => p.name === winner.name);

    // Calculate rotation
    const segmentAngle = 360 / participants.length;
    const spins = 8 + Math.floor(Math.random() * 5);
    const targetAbsolute = 360 - (actualIndex + 0.5) * segmentAngle;
    let delta = targetAbsolute - (rotation % 360);
    if (delta <= 0) delta += 360;
    const totalRotation = spins * 360 + delta;

    // Random duration 7-12 seconds
    const duration = 7 + Math.random() * 5;
    setSpinDuration(duration);

    // Prize name for announcements
    const prizeName = selectedPrize.label.split(' ').slice(1).join(' ');

    // === SOUND + VOICE SEQUENCE ===
    
    // 1. Opening announcement (excited, faster)
    const openingPhrases = [
      `Ladies and gentlemen! Time for the ${prizeName}!`,
      `Here we go! Drawing for ${prizeName}!`,
      `Get ready everyone! ${prizeName} is next!`
    ];
    speak(openingPhrases[Math.floor(Math.random() * openingPhrases.length)], 1.1, 1.1, 1);
    
    // 2. Start sound effect
    playStartSound();
    
    // 3. Start wheel rotation after brief pause
    setTimeout(() => {
      setRotation(prev => prev + totalRotation);
      
      // 4. Start tick sounds
      playDrumRoll(duration - 1);
    }, 800);
    
    // 5. Suspense phrases during spin (building tension)
    setTimeout(() => {
      speak("The wheel is spinning!", 1.0, 1.0, 0.9);
    }, 2000);
    
    setTimeout(() => {
      const suspensePhrases = [
        "Who will be the lucky one?",
        "Around and around it goes!",
        "The tension is building!"
      ];
      speak(suspensePhrases[Math.floor(Math.random() * suspensePhrases.length)], 0.95, 1.0, 0.9);
    }, 4000);
    
    setTimeout(() => {
      speak("Getting closer!", 1.0, 1.1, 1);
    }, (duration - 3) * 1000);
    
    // 6. Countdown beeps
    setTimeout(() => {
      playCountdown();
    }, (duration - 1.5) * 1000);
    
    // 7. The big reveal (slower, dramatic)
    setTimeout(() => {
      speak("And the winner is...", 0.8, 0.9, 1);
    }, (duration - 1) * 1000);

    // 8. Winner announcement + sound effect
    setTimeout(() => {
      playBigWin();
    }, duration * 1000);

    // Handle result after wheel stops
    setTimeout(() => {
      const prizeLabel = selectedPrize.label;

      const newParticipants = participants.map(p =>
        p.name === winner.name ? { ...p, won: true, prize: selectedPrize.name } : p
      );

      const newWinner: Winner = {
        name: winner.name,
        prize: prizeLabel,
        time: new Date().toLocaleString()
      };
      const newWinners = [...winners, newWinner];

      // Update prize config
      const newPrizeConfig = prizeConfig.map((pc, idx) =>
        idx === selectedPrizeIndex ? { ...pc, drawn: pc.drawn + 1 } : pc
      );

      setParticipants(newParticipants);
      setWinners(newWinners);
      setPrizeConfig(newPrizeConfig);
      setCurrentWinner({ name: winner.name, prize: prizeLabel });
      setShowFireworks(true);
      saveData(newParticipants, newWinners, newPrizeConfig);

      // 9. Announce winner name (excited, louder)
      setTimeout(() => {
        speak(`${winner.name}!`, 0.9, 1.2, 1);
      }, 300);
      
      // 10. Congratulations (celebratory)
      setTimeout(() => {
        const congratsPhrases = [
          "Congratulations! Give them a big round of applause!",
          "Amazing! What a lucky winner!",
          "Wonderful! Congratulations to our winner!"
        ];
        speak(congratsPhrases[Math.floor(Math.random() * congratsPhrases.length)], 1.0, 1.1, 1);
      }, 1500);

      setTimeout(() => setShowFireworks(false), 3000);
      setIsSpinning(false);
    }, duration * 1000 + 200);
  };

  const totalCount = participants.length;
  const remainCount = participants.filter(p => !p.won).length;

  if (isLoading) {
    return (
      <div className="loading-container">
        <div className="loading-text">Loading...</div>
      </div>
    );
  }

  return (
    <div className="lottery-container">
      <h1 className="main-title">🎯 Lucky Draw Wheel</h1>

      <div className={`grid-container ${leftPanelCollapsed ? 'left-collapsed' : ''} ${rightPanelCollapsed ? 'right-collapsed' : ''}`}>
        {/* Left - Participant Management */}
        <div className={`panel panel-left ${leftPanelCollapsed ? 'collapsed' : ''}`}>
          <button 
            className="panel-toggle panel-toggle-left"
            onClick={() => setLeftPanelCollapsed(!leftPanelCollapsed)}
            title={leftPanelCollapsed ? 'Expand' : 'Collapse'}
          >
            {leftPanelCollapsed ? '▶' : '◀'}
          </button>
          
          {leftPanelCollapsed ? (
            <div className="collapsed-content">
              <span className="collapsed-title">👥</span>
              <span className="collapsed-count">{totalCount}</span>
            </div>
          ) : (
            <>
          <h3 className="panel-title">👥 Participants</h3>

          {/* Excel Import Section */}
          <div className="excel-section">
            <div className="excel-title">📊 Excel Batch Import</div>
            <div className="excel-buttons">
              <button onClick={downloadTemplate} className="excel-btn">
                📥 Download Template
              </button>
              <label className="excel-btn">
                📤 Upload List
                <input 
                  type="file" 
                  accept=".xlsx,.xls" 
                  onChange={handleExcelUpload} 
                  className="hidden-input" 
                />
              </label>
            </div>
            <div className="excel-hint">
              Format: No. | Name (Name column auto-detected)
            </div>
          </div>

          <div className="input-row">
            <input
              type="text"
              value={nameInput}
              onChange={e => setNameInput(e.target.value)}
              onKeyPress={e => e.key === 'Enter' && addParticipant()}
              placeholder="Enter name"
              className="text-input"
            />
            <button onClick={addParticipant} className="btn-primary">
              Add
            </button>
          </div>

          <textarea
            value={batchInput}
            onChange={e => setBatchInput(e.target.value)}
            placeholder="Paste names here&#10;One per line, or separated by comma/space"
            className="textarea-input"
          />
          <button onClick={batchAdd} className="btn-secondary">
            Batch Add
          </button>

          <div className="participant-list">
            {participants.map((p, index) => (
              <div key={index} className={`participant-item ${p.won ? 'won' : ''}`}>
                <span>{index + 1}. {p.name} {p.won && '(Won)'}</span>
                <button onClick={() => removeParticipant(index)} className="delete-btn">
                  ×
                </button>
              </div>
            ))}
          </div>

          <div className="stats">
            <span>Total: <strong>{totalCount}</strong></span>
            <span>Remaining: <strong>{remainCount}</strong></span>
          </div>

          <div className="btn-group">
            <button onClick={clearAll} className="btn-small btn-danger">🗑️ Clear All</button>
            <button onClick={resetWinners} className="btn-small btn-export">🔄 Reset Draw</button>
          </div>
            </>
          )}
        </div>

        {/* Center - Wheel */}
        <div className="panel wheel-panel">
          <div className="wheel-container" style={{ width: leftPanelCollapsed || rightPanelCollapsed ? '480px' : '380px', height: leftPanelCollapsed || rightPanelCollapsed ? '480px' : '380px' }}>
            {/* Pointer */}
            <div className="pointer" />

            {/* Wheel SVG */}
            {(() => {
              const wheelSize = leftPanelCollapsed || rightPanelCollapsed ? 480 : 380;
              const center = wheelSize / 2;
              const outerRadius = wheelSize / 2 - 5;
              const innerRadius = outerRadius - 7;
              const segmentRadius = outerRadius - 10;
              
              return (
              <svg
                width={wheelSize}
                height={wheelSize}
                viewBox={`0 0 ${wheelSize} ${wheelSize}`}
                className="wheel-svg"
                style={{
                  transform: `rotate(${rotation}deg)`,
                  transition: isSpinning 
                    ? `transform ${spinDuration}s cubic-bezier(0.17, 0.67, 0.12, 0.99)` 
                    : 'none',
                }}
              >
                <circle cx={center} cy={center} r={outerRadius} fill="none" stroke="rgba(255,215,0,0.5)" strokeWidth="4" />
                <circle cx={center} cy={center} r={innerRadius} fill="none" stroke="rgba(255,255,255,0.2)" strokeWidth="8" />
                
                {participants.length === 0 ? (
                  <circle cx={center} cy={center} r={segmentRadius} fill="#4B5563" />
                ) : (
                  participants.map((p, index) => {
                    const total = participants.length;
                    const segmentAngle = 360 / total;
                    const startAngle = index * segmentAngle - 90;
                    const endAngle = startAngle + segmentAngle;
                    
                    const startRad = (startAngle * Math.PI) / 180;
                    const endRad = (endAngle * Math.PI) / 180;
                    
                    const radius = segmentRadius;
                    const cx = center;
                    const cy = center;
                    
                    const x1 = cx + radius * Math.cos(startRad);
                    const y1 = cy + radius * Math.sin(startRad);
                    const x2 = cx + radius * Math.cos(endRad);
                    const y2 = cy + radius * Math.sin(endRad);
                    
                    const largeArc = segmentAngle > 180 ? 1 : 0;
                    const pathD = `M ${cx} ${cy} L ${x1} ${y1} A ${radius} ${radius} 0 ${largeArc} 1 ${x2} ${y2} Z`;
                    
                    // Text position - along the radius, near the edge
                    const midAngle = startAngle + segmentAngle / 2;
                    const midRad = (midAngle * Math.PI) / 180;
                    
                    // Calculate text path along radius
                    const textStartRadius = radius * 0.92;
                    const textEndRadius = radius * 0.25;
                    const textStartX = cx + textStartRadius * Math.cos(midRad);
                    const textStartY = cy + textStartRadius * Math.sin(midRad);
                    const textEndX = cx + textEndRadius * Math.cos(midRad);
                    const textEndY = cy + textEndRadius * Math.sin(midRad);
                    
                    const textPathId = `textPath-${index}`;
                    const fontSize = total > 30 ? "9" : total > 20 ? "10" : total > 10 ? "12" : "14";
                    
                    return (
                      <g key={index}>
                        <path
                          d={pathD}
                          fill={COLORS[index % COLORS.length]}
                          stroke="rgba(255,255,255,0.3)"
                          strokeWidth="1"
                          style={{ filter: p.won ? 'grayscale(1) brightness(0.4)' : 'none' }}
                        />
                        {/* Define path for text to follow - from edge toward center */}
                        <defs>
                          <path
                            id={textPathId}
                            d={`M ${textStartX} ${textStartY} L ${textEndX} ${textEndY}`}
                          />
                        </defs>
                        <text
                          fill="white"
                          fontSize={fontSize}
                          fontWeight="600"
                          style={{
                            filter: p.won ? 'grayscale(1) brightness(0.4)' : 'none'
                          }}
                        >
                          <textPath
                            href={`#${textPathId}`}
                            startOffset="0%"
                            style={{ textShadow: '1px 1px 2px rgba(0,0,0,0.8)' }}
                          >
                            {p.name.length > 10 ? p.name.slice(0, 9) + '..' : p.name}
                          </textPath>
                        </text>
                      </g>
                    );
                  })
                )}
              </svg>
              );
            })()}

            {/* Center Circle */}
            <div 
              className="center-circle"
              style={{
                width: leftPanelCollapsed || rightPanelCollapsed ? '90px' : '70px',
                height: leftPanelCollapsed || rightPanelCollapsed ? '90px' : '70px',
                fontSize: leftPanelCollapsed || rightPanelCollapsed ? '40px' : '32px'
              }}
            >🎁</div>
          </div>

          {/* Winner Display */}
          <div className="winner-display">
            {isSpinning ? (
              <div className="spinning-text">
                <p className="winner-label">🎰 Spinning...</p>
                <p className="suspense-text">Who will be the lucky winner?</p>
              </div>
            ) : currentWinner ? (
              <div className="winner-reveal">
                <p className="winner-prize-label">{currentWinner.prize}</p>
                <p className="winner-announce">🎉 THE WINNER IS 🎉</p>
                <p className="winner-name">{currentWinner.name}</p>
              </div>
            ) : (
              <p className="winner-label">Click the button below to start</p>
            )}
          </div>

          <button
            onClick={spin}
            disabled={isSpinning || remainCount === 0 || !getSelectedPrize()}
            className="spin-btn"
          >
            {getSelectedPrize() 
              ? `🎰 SPIN for ${getSelectedPrize()?.label.split(' ').slice(1).join(' ')}`
              : '🎰 Select a Prize'}
          </button>

          {/* Sound Control Panel */}
          <div className="sound-control-panel">
            <div className="sound-control-header">
              <span className="sound-control-title">🎵 Sound & Voice</span>
            </div>
            
            <div className="sound-control-options">
              <div className="sound-toggle-row">
                <button 
                  onClick={() => setSoundEnabled(!soundEnabled)}
                  className={`sound-toggle-btn ${soundEnabled ? 'on' : 'off'}`}
                >
                  {soundEnabled ? '🔔 Effects ON' : '🔕 Effects OFF'}
                </button>
                <button 
                  onClick={() => setVoiceEnabled(!voiceEnabled)}
                  className={`sound-toggle-btn ${voiceEnabled ? 'on' : 'off'}`}
                >
                  {voiceEnabled ? '🎙️ Voice ON' : '🔇 Voice OFF'}
                </button>
              </div>
              
              <div className="sound-control-row">
                <label>Volume:</label>
                <input
                  type="range"
                  min="0"
                  max="1"
                  step="0.1"
                  value={soundVolume}
                  onChange={e => setSoundVolume(parseFloat(e.target.value))}
                  className="sound-slider"
                />
                <span className="sound-value">{Math.round(soundVolume * 100)}%</span>
              </div>
              
              {voiceEnabled && availableVoices.length > 0 && (
                <div className="sound-control-row">
                  <label>Voice:</label>
                  <select
                    value={selectedVoice?.name || ''}
                    onChange={e => {
                      const voice = availableVoices.find(v => v.name === e.target.value);
                      setSelectedVoice(voice || null);
                    }}
                    className="voice-select"
                  >
                    {availableVoices.filter(v => v.lang.startsWith('en')).slice(0, 10).map(voice => (
                      <option key={voice.name} value={voice.name}>
                        {voice.name.length > 20 ? voice.name.slice(0, 20) + '...' : voice.name}
                      </option>
                    ))}
                  </select>
                </div>
              )}
              
              <button 
                onClick={() => {
                  playStartSound();
                  setTimeout(() => playTick(1, 800), 300);
                  setTimeout(() => playTick(1, 700), 450);
                  setTimeout(() => playFanfare(), 600);
                  setTimeout(() => speak("Testing! One, two, three!", 1, 1.1, 1), 100);
                }}
                className="sound-test-btn"
              >
                🎰 Test Sound & Voice
              </button>
            </div>
          </div>
        </div>

        {/* Right - Winners List */}
        <div className={`panel panel-right ${rightPanelCollapsed ? 'collapsed' : ''}`}>
          <button 
            className="panel-toggle panel-toggle-right"
            onClick={() => setRightPanelCollapsed(!rightPanelCollapsed)}
            title={rightPanelCollapsed ? 'Expand' : 'Collapse'}
          >
            {rightPanelCollapsed ? '◀' : '▶'}
          </button>
          
          {rightPanelCollapsed ? (
            <div className="collapsed-content">
              <span className="collapsed-title">🏆</span>
              <span className="collapsed-count">{winners.length}</span>
            </div>
          ) : (
            <>
          <h3 className="panel-title">🏆 Prize & Winners</h3>

          {/* Prize Configuration */}
          <div className="prize-config-section">
            <div className="prize-config-title">Prize Settings</div>
            {prizeConfig.map((prize, index) => (
              <div key={prize.name} className="prize-config-row">
                <span className="prize-config-label">{prize.label}</span>
                <div className="prize-config-input-group">
                  <input
                    type="number"
                    min="0"
                    max="100"
                    value={prize.count}
                    onChange={e => {
                      const newConfig = [...prizeConfig];
                      newConfig[index].count = Math.max(0, parseInt(e.target.value) || 0);
                      setPrizeConfig(newConfig);
                      setConfigSaved(false);
                    }}
                    className="prize-config-input"
                  />
                  <span className="prize-config-drawn">
                    ({prize.drawn}/{prize.count})
                  </span>
                </div>
              </div>
            ))}
            <div className="prize-config-footer">
              <div className="prize-config-summary">
                Total: {prizeConfig.reduce((sum, p) => sum + p.count, 0)} | 
                Drawn: {prizeConfig.reduce((sum, p) => sum + p.drawn, 0)}
              </div>
              <button 
                onClick={() => {
                  saveData(participants, winners, prizeConfig);
                  setConfigSaved(true);
                }}
                className={`btn-save ${configSaved ? 'saved' : ''}`}
              >
                {configSaved ? '✓ Saved' : '💾 Save'}
              </button>
            </div>
          </div>

          {/* Prize Selection */}
          <div className="prize-selection-section">
            <div className="prize-selection-title">Select Prize to Draw</div>
            <div className="prize-selection-buttons">
              {prizeConfig.map((prize, index) => {
                const isCompleted = prize.count === 0 || prize.drawn >= prize.count;
                const isSelected = selectedPrizeIndex === index;
                return (
                  <button
                    key={prize.name}
                    onClick={() => !isCompleted && setSelectedPrizeIndex(index)}
                    className={`prize-btn ${isSelected ? 'selected' : ''} ${isCompleted ? 'completed' : ''}`}
                    disabled={isCompleted}
                  >
                    <span className="prize-btn-label">{prize.label}</span>
                    <span className="prize-btn-count">
                      {isCompleted ? '✓ Done' : `${prize.drawn}/${prize.count}`}
                    </span>
                  </button>
                );
              })}
            </div>
          </div>

          {/* Winner List */}
          <div className="winner-list-title">Winner List ({winners.length})</div>
          <div className="winner-list">
            {winners.length === 0 ? (
              <div className="no-winners">No winners yet</div>
            ) : (
              winners.map((w, index) => (
                <div key={index} className="winner-item">
                  <div>
                    <div className="winner-prize">{w.prize}</div>
                    <div className="winner-item-name">{w.name}</div>
                  </div>
                  <div className="winner-time">{w.time}</div>
                </div>
              ))
            )}
          </div>

          <div className="export-buttons">
            <button onClick={exportWinners} className="btn-export-winner btn-txt">
              📋 Export Winners (TXT)
            </button>
            <button onClick={exportWinnersExcel} className="btn-export-winner btn-excel">
              📊 Export Winners (Excel)
            </button>
          </div>
            </>
          )}
        </div>
      </div>

      {/* Fireworks Effect */}
      {showFireworks && (
        <div className="fireworks">
          {/* Large firework bursts */}
          {[...Array(30)].map((_, i) => (
            <div
              key={`fw-${i}`}
              className="firework"
              style={{
                left: `${Math.random() * 100}%`,
                top: `${Math.random() * 100}%`,
                background: COLORS[Math.floor(Math.random() * COLORS.length)],
                animationDelay: `${Math.random() * 0.8}s`,
                width: `${10 + Math.random() * 15}px`,
                height: `${10 + Math.random() * 15}px`,
              }}
            />
          ))}
          {/* Small sparkles */}
          {[...Array(50)].map((_, i) => (
            <div
              key={`sp-${i}`}
              className="sparkle"
              style={{
                left: `${Math.random() * 100}%`,
                top: `${Math.random() * 100}%`,
                animationDelay: `${Math.random() * 1.2}s`,
                background: Math.random() > 0.5 ? '#FFD700' : '#FFFFFF',
              }}
            />
          ))}
        </div>
      )}
    </div>
  );
}