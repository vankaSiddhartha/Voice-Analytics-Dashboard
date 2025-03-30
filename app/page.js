"use client"
// pages/index.js
import { useState, useEffect, useRef } from 'react';
import * as XLSX from 'xlsx';
import { Bar, Line, Pie } from 'react-chartjs-2';
import {
  Chart as ChartJS,
  CategoryScale,
  LinearScale,
  BarElement,
  LineElement,
  PointElement,
  ArcElement,
  Title,
  Tooltip,
  Legend,
} from 'chart.js';

// Register Chart.js components
ChartJS.register(
  CategoryScale,
  LinearScale,
  BarElement,
  LineElement,
  PointElement,
  ArcElement,
  Title,
  Tooltip,
  Legend
);

export default function Home() {
  const [data, setData] = useState([]);
  const [filteredData, setFilteredData] = useState([]);
  const [countries, setCountries] = useState([]);
  const [listening, setListening] = useState(false);
  const [transcript, setTranscript] = useState('');
  const [chartType, setChartType] = useState('bar');
  const [metric, setMetric] = useState('Confirmed');
  const [selectedCountries, setSelectedCountries] = useState([]);
  const [status, setStatus] = useState('Upload your XLSX file to begin');
  const [darkMode, setDarkMode] = useState(false);
  
  const recognitionRef = useRef(null);

  useEffect(() => {
    // Initialize speech recognition when component mounts
    if ('webkitSpeechRecognition' in window || 'SpeechRecognition' in window) {
      const SpeechRecognition = window.SpeechRecognition || window.webkitSpeechRecognition;
      recognitionRef.current = new SpeechRecognition();
      recognitionRef.current.continuous = true;
      recognitionRef.current.interimResults = true;
      
      recognitionRef.current.onresult = (event) => {
        const transcript = Array.from(event.results)
          .map(result => result[0])
          .map(result => result.transcript)
          .join('');
        
        setTranscript(transcript);
        processVoiceCommand(transcript);
      };
      
      recognitionRef.current.onerror = (event) => {
        console.error('Speech recognition error', event.error);
        setStatus(`Error: ${event.error}`);
      };
    } else {
      setStatus('Speech recognition not supported in your browser');
    }
    
    return () => {
      if (recognitionRef.current) {
        recognitionRef.current.stop();
      }
    };
  }, [data, countries]);

  const processVoiceCommand = (command) => {
    // Convert command to lowercase for easier matching
    const cmd = command.toLowerCase();
    
    // Look for country names in the command
    const countryMatches = countries.filter(country => 
      cmd.includes(country.toLowerCase())
    );
    
    if (countryMatches.length > 0) {
      setStatus(`Found countries: ${countryMatches.join(', ')}`);
      setSelectedCountries(countryMatches);
      updateFilteredData(countryMatches);
    }
    
    // Check for chart type commands
    if (cmd.includes('bar chart') || cmd.includes('bar graph')) {
      setChartType('bar');
      setStatus('Showing bar chart');
    } else if (cmd.includes('line chart') || cmd.includes('line graph')) {
      setChartType('line');
      setStatus('Showing line chart');
    } else if (cmd.includes('pie chart')) {
      setChartType('pie');
      setStatus('Showing pie chart');
    }
    
    // Check for metric commands
    if (cmd.includes('deaths') || cmd.includes('death')) {
      setMetric('Deaths');
      setStatus('Showing death data');
    } else if (cmd.includes('recovered') || cmd.includes('recovery')) {
      setMetric('Recovered');
      setStatus('Showing recovery data');
    } else if (cmd.includes('active') || cmd.includes('active cases')) {
      setMetric('Active');
      setStatus('Showing active cases');
    } else if (cmd.includes('confirmed') || cmd.includes('total cases')) {
      setMetric('Confirmed');
      setStatus('Showing confirmed cases');
    } else if (cmd.includes('new cases')) {
      setMetric('New cases');
      setStatus('Showing new cases');
    }
    
    // Handle dark mode toggle
    if (cmd.includes('dark mode') || cmd.includes('night mode')) {
      setDarkMode(true);
      setStatus('Dark mode activated');
    } else if (cmd.includes('light mode') || cmd.includes('day mode')) {
      setDarkMode(false);
      setStatus('Light mode activated');
    }
    
    // Handle comparison commands
    if (cmd.includes('compare') || cmd.includes('comparison')) {
      setStatus('Showing comparison between selected countries');
    }
    
    // Handle sorting commands
    if (cmd.includes('sort by') || cmd.includes('order by')) {
      if (cmd.includes('highest') || cmd.includes('most')) {
        const sortedData = [...filteredData].sort((a, b) => b[metric] - a[metric]);
        setFilteredData(sortedData.slice(0, 10));
        setStatus(`Showing countries with highest ${metric}`);
      } else if (cmd.includes('lowest') || cmd.includes('least')) {
        const sortedData = [...filteredData].sort((a, b) => a[metric] - b[metric]);
        setFilteredData(sortedData.slice(0, 10));
        setStatus(`Showing countries with lowest ${metric}`);
      }
    }
    
    // Handle top N commands
    const topMatch = cmd.match(/top (\d+)/);
    if (topMatch && topMatch[1]) {
      const topN = parseInt(topMatch[1]);
      const sortedData = [...data].sort((a, b) => b[metric] - a[metric]);
      setFilteredData(sortedData.slice(0, topN));
      setStatus(`Showing top ${topN} countries by ${metric}`);
    }
  };

  const toggleListening = () => {
    if (listening) {
      recognitionRef.current.stop();
      setListening(false);
      setStatus('Voice recognition stopped');
    } else {
      recognitionRef.current.start();
      setListening(true);
      setStatus('Listening... Speak a command');
    }
  };

  const handleFileUpload = (e) => {
    const file = e.target.files[0];
    const reader = new FileReader();
    
    reader.onload = (evt) => {
      try {
        const binaryData = evt.target.result;
        const workbook = XLSX.read(binaryData, { type: 'binary' });
        const worksheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[worksheetName];
        const jsonData = XLSX.utils.sheet_to_json(worksheet);
        
        setData(jsonData);
        setFilteredData(jsonData);
        
        // Extract unique countries
        const uniqueCountries = [...new Set(jsonData.map(item => item['Country/Region']))];
        setCountries(uniqueCountries);
        
        setStatus(`Loaded ${jsonData.length} records from ${uniqueCountries.length} countries`);
      } catch (error) {
        console.error('Error processing file:', error);
        setStatus('Error processing file. Please check the format.');
      }
    };
    
    reader.onerror = () => {
      setStatus('Error reading file');
    };
    
    if (file) {
      reader.readAsBinaryString(file);
    }
  };

  const updateFilteredData = (selectedCountries) => {
    if (selectedCountries.length === 0) {
      setFilteredData(data);
    } else {
      const filtered = data.filter(item => 
        selectedCountries.includes(item['Country/Region'])
      );
      setFilteredData(filtered);
    }
  };

  const renderChart = () => {
    if (filteredData.length === 0) return null;

    // Generate a dynamic color palette based on the number of data points
    const generateColors = (count) => {
      const colors = [];
      for (let i = 0; i < count; i++) {
        const hue = (i * 137.5) % 360; // Use golden ratio to distribute colors
        colors.push(`hsla(${hue}, 75%, 60%, 0.7)`);
      }
      return colors;
    };

    const backgroundColors = generateColors(filteredData.length);
    const borderColors = backgroundColors.map(color => color.replace('0.7', '1'));

    const chartData = {
      labels: filteredData.map(item => item['Country/Region']),
      datasets: [
        {
          label: metric,
          data: filteredData.map(item => item[metric] || 0),
          backgroundColor: backgroundColors,
          borderColor: borderColors,
          borderWidth: 2,
        },
      ],
    };

    const options = {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: {
          position: 'top',
          labels: {
            font: {
              family: "'Poppins', sans-serif",
              size: 14
            },
            color: darkMode ? '#ffffff' : '#333333',
          }
        },
        title: {
          display: true,
          text: `COVID-19 ${metric} by Country`,
          font: {
            family: "'Poppins', sans-serif",
            size: 18,
            weight: 'bold'
          },
          color: darkMode ? '#ffffff' : '#333333',
          padding: {
            top: 10,
            bottom: 20
          }
        },
        tooltip: {
          backgroundColor: darkMode ? 'rgba(30, 30, 30, 0.8)' : 'rgba(255, 255, 255, 0.8)',
          titleColor: darkMode ? '#ffffff' : '#333333',
          bodyColor: darkMode ? '#ffffff' : '#333333',
          borderColor: darkMode ? '#555555' : '#dddddd',
          borderWidth: 1,
          cornerRadius: 8,
          padding: 12,
          boxPadding: 4,
          usePointStyle: true,
        }
      },
      scales: {
        x: {
          ticks: {
            color: darkMode ? '#cccccc' : '#555555',
            font: {
              family: "'Poppins', sans-serif"
            }
          },
          grid: {
            color: darkMode ? 'rgba(255, 255, 255, 0.1)' : 'rgba(0, 0, 0, 0.1)'
          }
        },
        y: {
          ticks: {
            color: darkMode ? '#cccccc' : '#555555',
            font: {
              family: "'Poppins', sans-serif"
            }
          },
          grid: {
            color: darkMode ? 'rgba(255, 255, 255, 0.1)' : 'rgba(0, 0, 0, 0.1)'
          }
        }
      }
    };

    switch (chartType) {
      case 'bar':
        return <Bar data={chartData} options={options} />;
      case 'line':
        return <Line data={chartData} options={options} />;
      case 'pie':
        return <Pie data={chartData} options={options} />;
      default:
        return <Bar data={chartData} options={options} />;
    }
  };

  // Dynamically set styles based on dark mode
  const styles = {
    container: {
      fontFamily: "'Poppins', 'Roboto', sans-serif",
      margin: '0 auto',
      padding: '20px',
      maxWidth: '1200px',
      background: darkMode ? '#1f2937' : '#f8fafc',
      color: darkMode ? '#f3f4f6' : '#1e293b',
      minHeight: '100vh',
      transition: 'all 0.3s ease',
      borderRadius: '10px',
      boxShadow: darkMode ? '0 0 15px rgba(0, 0, 0, 0.3)' : '0 0 15px rgba(0, 0, 0, 0.1)'
    },
    header: {
      textAlign: 'center',
      borderBottom: darkMode ? '2px solid #4f46e5' : '2px solid #4338ca',
      marginBottom: '30px',
      paddingBottom: '15px',
      position: 'relative'
    },
    title: {
      fontSize: '36px',
      fontWeight: 'bold',
      margin: '0 0 10px 0',
      background: '-webkit-linear-gradient(45deg, #4338ca, #ec4899)',
      WebkitBackgroundClip: 'text',
      WebkitTextFillColor: 'transparent',
      textShadow: darkMode ? '0 0 8px rgba(79, 70, 229, 0.4)' : 'none',
    },
    statusCard: {
      background: darkMode ? '#374151' : 'white',
      borderRadius: '12px',
      padding: '20px',
      marginBottom: '30px',
      boxShadow: darkMode 
        ? '0 4px 6px -1px rgba(0, 0, 0, 0.3), 0 2px 4px -2px rgba(0, 0, 0, 0.2)' 
        : '0 4px 6px -1px rgba(0, 0, 0, 0.1), 0 2px 4px -2px rgba(0, 0, 0, 0.05)',
      position: 'relative',
      overflow: 'hidden'
    },
    statusLabel: {
      fontSize: '18px',
      fontWeight: '600',
      marginBottom: '10px',
      color: darkMode ? '#93c5fd' : '#3b82f6'
    },
    transcript: {
      fontSize: '16px',
      padding: '15px',
      background: darkMode ? '#1f2937' : '#f1f5f9',
      borderRadius: '8px',
      marginBottom: '15px',
      borderLeft: darkMode ? '4px solid #4f46e5' : '4px solid #4338ca',
      maxHeight: '100px',
      overflowY: 'auto'
    },
    buttonContainer: {
      display: 'flex',
      gap: '15px',
      marginBottom: '20px',
      flexWrap: 'wrap'
    },
    listeningButton: {
      padding: '12px 20px',
      borderRadius: '8px',
      border: 'none',
      fontWeight: '600',
      fontSize: '16px',
      cursor: 'pointer',
      display: 'flex',
      alignItems: 'center',
      justifyContent: 'center',
      gap: '8px',
      transition: 'all 0.2s ease',
      background: listening 
        ? 'linear-gradient(45deg, #ef4444, #f43f5e)' 
        : 'linear-gradient(45deg, #22c55e, #10b981)',
      color: 'white',
      boxShadow: listening
        ? '0 4px 6px -1px rgba(239, 68, 68, 0.5), 0 2px 4px -2px rgba(239, 68, 68, 0.4)'
        : '0 4px 6px -1px rgba(34, 197, 94, 0.5), 0 2px 4px -2px rgba(34, 197, 94, 0.4)',
    },
    uploadButton: {
      position: 'relative',
      padding: '12px 20px',
      borderRadius: '8px',
      border: 'none',
      fontWeight: '600',
      fontSize: '16px',
      cursor: 'pointer',
      background: 'linear-gradient(45deg, #6366f1, #4f46e5)',
      color: 'white',
      boxShadow: '0 4px 6px -1px rgba(99, 102, 241, 0.5), 0 2px 4px -2px rgba(99, 102, 241, 0.4)',
      display: 'flex',
      alignItems: 'center',
      justifyContent: 'center',
      gap: '8px',
      transition: 'all 0.2s ease',
    },
    darkModeToggle: {
      position: 'relative',
      padding: '12px 20px',
      borderRadius: '8px',
      border: 'none',
      fontWeight: '600',
      fontSize: '16px',
      cursor: 'pointer',
      background: darkMode 
        ? 'linear-gradient(45deg, #fbbf24, #f59e0b)' 
        : 'linear-gradient(45deg, #1e40af, #1e3a8a)',
      color: 'white',
      boxShadow: darkMode
        ? '0 4px 6px -1px rgba(251, 191, 36, 0.5), 0 2px 4px -2px rgba(251, 191, 36, 0.4)'
        : '0 4px 6px -1px rgba(30, 64, 175, 0.5), 0 2px 4px -2px rgba(30, 64, 175, 0.4)',
      display: 'flex',
      alignItems: 'center',
      justifyContent: 'center',
      gap: '8px',
      transition: 'all 0.2s ease',
    },
    fileInput: {
      position: 'absolute',
      inset: 0,
      opacity: 0,
      cursor: 'pointer',
      width: '100%',
      height: '100%'
    },
    commandsCard: {
      background: darkMode ? '#374151' : 'white',
      borderRadius: '12px',
      padding: '20px',
      marginBottom: '30px',
      boxShadow: darkMode 
        ? '0 4px 6px -1px rgba(0, 0, 0, 0.3), 0 2px 4px -2px rgba(0, 0, 0, 0.2)' 
        : '0 4px 6px -1px rgba(0, 0, 0, 0.1), 0 2px 4px -2px rgba(0, 0, 0, 0.05)',
    },
    commandsTitle: {
      fontSize: '20px',
      fontWeight: '600',
      marginBottom: '15px',
      color: darkMode ? '#93c5fd' : '#3b82f6',
      display: 'flex',
      alignItems: 'center',
      gap: '8px'
    },
    commandsList: {
      display: 'grid',
      gridTemplateColumns: 'repeat(auto-fill, minmax(250px, 1fr))',
      gap: '15px',
      padding: 0,
      margin: 0,
      listStyle: 'none'
    },
    commandItem: {
      padding: '10px 15px',
      background: darkMode ? '#1f2937' : '#f1f5f9',
      borderRadius: '8px',
      fontSize: '14px',
      fontWeight: '500',
      position: 'relative',
      paddingLeft: '25px',
      border: darkMode ? '1px solid #4b5563' : '1px solid #e2e8f0'
    },
    commandBullet: {
      position: 'absolute',
      left: '10px',
      top: '50%',
      transform: 'translateY(-50%)',
      width: '8px',
      height: '8px',
      borderRadius: '50%',
      background: 'linear-gradient(45deg, #4f46e5, #6366f1)'
    },
    visSection: {
      background: darkMode ? '#374151' : 'white',
      borderRadius: '12px',
      padding: '20px',
      marginBottom: '30px',
      boxShadow: darkMode 
        ? '0 4px 6px -1px rgba(0, 0, 0, 0.3), 0 2px 4px -2px rgba(0, 0, 0, 0.2)' 
        : '0 4px 6px -1px rgba(0, 0, 0, 0.1), 0 2px 4px -2px rgba(0, 0, 0, 0.05)',
    },
    visTitle: {
      fontSize: '20px',
      fontWeight: '600',
      marginBottom: '15px',
      color: darkMode ? '#93c5fd' : '#3b82f6',
      display: 'flex',
      alignItems: 'center',
      gap: '8px'
    },
    controlsRow: {
      display: 'flex',
      gap: '15px',
      marginBottom: '20px',
      flexWrap: 'wrap'
    },
    selectBox: {
      padding: '10px 15px',
      borderRadius: '8px',
      border: darkMode ? '1px solid #4b5563' : '1px solid #e2e8f0',
      background: darkMode ? '#1f2937' : 'white',
      color: darkMode ? '#f3f4f6' : '#1e293b',
      fontFamily: "'Poppins', sans-serif",
      fontSize: '14px',
      fontWeight: '500',
      cursor: 'pointer',
      outline: 'none',
      boxShadow: '0 1px 2px rgba(0, 0, 0, 0.05)',
      transition: 'all 0.2s ease'
    },
    chartContainer: {
      height: '400px',
      position: 'relative',
      marginBottom: '30px',
      background: darkMode ? '#1f2937' : '#f8fafc',
      borderRadius: '8px',
      padding: '20px',
      border: darkMode ? '1px solid #4b5563' : '1px solid #e2e8f0'
    },
    tableSection: {
      background: darkMode ? '#374151' : 'white',
      borderRadius: '12px',
      padding: '20px',
      boxShadow: darkMode 
        ? '0 4px 6px -1px rgba(0, 0, 0, 0.3), 0 2px 4px -2px rgba(0, 0, 0, 0.2)' 
        : '0 4px 6px -1px rgba(0, 0, 0, 0.1), 0 2px 4px -2px rgba(0, 0, 0, 0.05)',
      overflowX: 'auto'
    },
    tableTitle: {
      fontSize: '20px',
      fontWeight: '600',
      marginBottom: '15px',
      color: darkMode ? '#93c5fd' : '#3b82f6',
    },
    table: {
      width: '100%',
      borderCollapse: 'separate',
      borderSpacing: 0,
      borderRadius: '8px',
      overflow: 'hidden',
      border: darkMode ? '1px solid #4b5563' : '1px solid #e2e8f0'
    },
    tableHeader: {
      background: darkMode ? '#1f2937' : '#f1f5f9',
      color: darkMode ? '#f3f4f6' : '#1e293b',
      textAlign: 'left',
      fontWeight: '600',
      fontSize: '14px',
      textTransform: 'uppercase',
      letterSpacing: '0.05em'
    },
    tableHeaderCell: {
      padding: '12px 16px',
      borderBottom: darkMode ? '1px solid #4b5563' : '1px solid #e2e8f0'
    },
    tableRow: {
      transition: 'background-color 0.2s ease'
    },
    tableRowEven: {
      background: darkMode ? '#283548' : '#f8fafc'
    },
    tableRowOdd: {
      background: darkMode ? '#1f2937' : 'white'
    },
    tableCell: {
      padding: '12px 16px',
      borderBottom: darkMode ? '1px solid #374151' : '1px solid #e2e8f0',
      fontSize: '14px'
    },
    tableCellNumeric: {
      textAlign: 'right',
      fontVariantNumeric: 'tabular-nums'
    },
    loader: {
      display: 'flex',
      justifyContent: 'center',
      alignItems: 'center',
      height: '200px',
      width: '100%'
    },
    footer: {
      textAlign: 'center',
      marginTop: '40px',
      paddingTop: '20px',
      borderTop: darkMode ? '1px solid #4b5563' : '1px solid #e2e8f0',
      fontSize: '14px',
      color: darkMode ? '#9ca3af' : '#6b7280'
    },
    pulseDot: {
      position: 'absolute',
      top: '15px',
      right: '15px',
      width: '12px',
      height: '12px',
      borderRadius: '50%',
      background: listening ? '#ef4444' : '#22c55e',
      boxShadow: listening 
        ? '0 0 0 0 rgba(239, 68, 68, 1)' 
        : '0 0 0 0 rgba(34, 197, 94, 1)',
      animation: listening 
        ? 'pulse-red 2s infinite' 
        : 'pulse-green 2s infinite',
    },
    keyStylesheet: `
      @import url('https://fonts.googleapis.com/css2?family=Poppins:wght@300;400;500;600;700&display=swap');
      
      @keyframes pulse-red {
        0% {
          box-shadow: 0 0 0 0 rgba(239, 68, 68, 0.7);
        }
        70% {
          box-shadow: 0 0 0 10px rgba(239, 68, 68, 0);
        }
        100% {
          box-shadow: 0 0 0 0 rgba(239, 68, 68, 0);
        }
      }
      
      @keyframes pulse-green {
        0% {
          box-shadow: 0 0 0 0 rgba(34, 197, 94, 0.7);
        }
        70% {
          box-shadow: 0 0 0 10px rgba(34, 197, 94, 0);
        }
        100% {
          box-shadow: 0 0 0 0 rgba(34, 197, 94, 0);
        }
      }
      
      body {
        margin: 0;
        padding: 0;
        background: ${darkMode ? '#111827' : '#f1f5f9'};
        transition: background 0.3s ease;
      }
      
      ::-webkit-scrollbar {
        width: 8px;
        height: 8px;
      }
      
      ::-webkit-scrollbar-track {
        background: ${darkMode ? '#1f2937' : '#f1f5f9'};
        border-radius: 10px;
      }
      
      ::-webkit-scrollbar-thumb {
        background: ${darkMode ? '#4b5563' : '#cbd5e1'};
        border-radius: 10px;
      }
      
      ::-webkit-scrollbar-thumb:hover {
        background: ${darkMode ? '#6b7280' : '#94a3b8'};
      }
    `
  };

  return (
    <div style={styles.container}>
      <style dangerouslySetInnerHTML={{ __html: styles.keyStylesheet }} />
      
      <header style={styles.header}>
        <h1 style={styles.title}> Voice Analytics Dashboard</h1>
      </header>
      
      <div style={styles.statusCard}>
        <div style={styles.pulseDot}></div>
        <h2 style={styles.statusLabel}>Current Status: {status}</h2>
        <div style={styles.transcript}>
          <p style={{ margin: 0 }}>{transcript || "No voice input detected yet..."}</p>
        </div>
        
        <div style={styles.buttonContainer}>
          <button
            onClick={toggleListening}
            style={styles.listeningButton}
          >
            {listening ? '🛑 Stop Listening' : '🎤 Start Listening'}
          </button>
          
          <div style={styles.uploadButton}>
            📊 Upload XLSX File
            <input
              type="file"
              accept=".xlsx,.xls"
              onChange={handleFileUpload}
              style={styles.fileInput}
            />
          </div>
          
          <button 
            onClick={() => setDarkMode(!darkMode)}
            style={styles.darkModeToggle}
          >
            {darkMode ? '☀️ Light Mode' : '🌙 Dark Mode'}
          </button>
        </div>
      </div>
      
      <div style={styles.commandsCard}>
        <h2 style={styles.commandsTitle}>
          🗣️ Voice Commands Examples:
        </h2>
        <ul style={styles.commandsList}>
          <li style={styles.commandItem}>
            <span style={styles.commandBullet}></span>
            "Show bar chart for United States and Italy"
          </li>
          <li style={styles.commandItem}>
            <span style={styles.commandBullet}></span>
            "Compare deaths in China, Spain, and France"
          </li>
          <li style={styles.commandItem}>
            <span style={styles.commandBullet}></span>
            "Show pie chart of confirmed cases"
          </li>
          <li style={styles.commandItem}>
            <span style={styles.commandBullet}></span>
            "Sort by highest deaths"
          </li>
          <li style={styles.commandItem}>
            <span style={styles.commandBullet}></span>
            "Show top 5 countries by recovered"
          </li>
          <li style={styles.commandItem}>
            <span style={styles.commandBullet}></span>
            "Switch to dark mode" / "Switch to light mode"
          </li>
        </ul>
      </div>
      
      <div style={styles.visSection}>
        <h2 style={styles.visTitle}>
          📈 Data Visualization
        </h2>
        <div style={styles.controlsRow}>
          <select
            value={chartType}
            onChange={(e) => setChartType(e.target.value)}
            style={styles.selectBox}
          >


              <option value="bar">Bar Chart</option>
            <option value="line">Line Chart</option>
            <option value="pie">Pie Chart</option>
          </select>
          
          <select
            value={metric}
            onChange={(e) => setMetric(e.target.value)}
            style={styles.selectBox}
          >
            <option value="Confirmed">Confirmed Cases</option>
            <option value="Deaths">Deaths</option>
            <option value="Recovered">Recovered</option>
            <option value="Active">Active Cases</option>
            <option value="New cases">New Cases</option>
          </select>
          
          <select
            value={selectedCountries.length === 0 ? 'all' : 'selected'}
            onChange={(e) => {
              if (e.target.value === 'all') {
                setSelectedCountries([]);
                updateFilteredData([]);
              }
            }}
            style={styles.selectBox}
          >
            <option value="all">All Countries</option>
            <option value="selected">Selected Countries</option>
          </select>
        </div>
        
        <div style={styles.chartContainer}>
          {data.length === 0 ? (
            <div style={styles.loader}>
              <p>Upload data to visualize</p>
            </div>
          ) : (
            renderChart()
          )}
        </div>
      </div>
      
      {filteredData.length > 0 && (
        <div style={styles.tableSection}>
          <h2 style={styles.tableTitle}>Data Summary</h2>
          <table style={styles.table}>
            <thead>
              <tr style={styles.tableHeader}>
                <th style={styles.tableHeaderCell}>Country/Region</th>
                <th style={styles.tableHeaderCell}>Confirmed</th>
                <th style={styles.tableHeaderCell}>Deaths</th>
                <th style={styles.tableHeaderCell}>Recovered</th>
                <th style={styles.tableHeaderCell}>Active</th>
              </tr>
            </thead>
            <tbody>
              {filteredData.slice(0, 10).map((item, index) => (
                <tr
                  key={item['Country/Region']}
                  style={{
                    ...styles.tableRow,
                    ...(index % 2 === 0 ? styles.tableRowEven : styles.tableRowOdd)
                  }}
                >
                  <td style={styles.tableCell}>{item['Country/Region']}</td>
                  <td style={{...styles.tableCell, ...styles.tableCellNumeric}}>
                    {(item['Confirmed'] || 0).toLocaleString()}
                  </td>
                  <td style={{...styles.tableCell, ...styles.tableCellNumeric}}>
                    {(item['Deaths'] || 0).toLocaleString()}
                  </td>
                  <td style={{...styles.tableCell, ...styles.tableCellNumeric}}>
                    {(item['Recovered'] || 0).toLocaleString()}
                  </td>
                  <td style={{...styles.tableCell, ...styles.tableCellNumeric}}>
                    {(item['Active'] || 0).toLocaleString()}
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}
      
      <footer style={styles.footer}>
        <p> Voice Analytics Dashboard © {new Date().getFullYear()}</p>
        <p>Upload XLSX files with COVID-19 data to visualize and analyze using voice commands</p>
      </footer>
    </div>
  );
}